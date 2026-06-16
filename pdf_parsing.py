"""
pdf_parsing.py

Convert PDF with visual track changes (underline = insertion, strikethrough = deletion)
into Markdown with <ins>/<del> HTML tags.

Detection layers (tried in priority order per span):
  1. char_flags  — PyMuPDF MuPDF-level detection (bit 0=strikeout, bit 1=underline)
  2. Annotations — PDF_ANNOT_STRIKE_OUT / PDF_ANNOT_UNDERLINE
  3. Line-art    — Horizontal lines/thin rects from page vector graphics
  4. OCR+OpenCV  — For scanned pages: OCR text + morphological line detection

Requirements:
    pip install pymupdf opencv-python-headless surya-ocr

    Surya is the default OCR engine (best text accuracy).
    Alternative engines:
      - docTR:     pip install "python-doctr[torch]"
      - Tesseract: brew install tesseract  (macOS) / sudo apt install tesseract-ocr (Linux)

    pymupdf4llm flow (--flow pymupdf4llm):
      pip install pymupdf4llm

Usage:
    python pdf_parsing.py REM-CTNF-NOA.pdf [output.md]
    python pdf_parsing.py input.pdf --ocr-engine surya
    python pdf_parsing.py input.pdf --ocr-engine doctr
    python pdf_parsing.py REM-CTNF-NOA.pdf --ocr-engine tesseract

    # pymupdf4llm flow — force-OCR entire document, output raw markdown
    python pdf_parsing.py input.pdf --flow pymupdf4llm
    python pdf_parsing.py input.pdf --flow pymupdf4llm output.md

    # As a library
    from pdf_parsing import convert
    md = convert("input.pdf", output_md="out.md")
    md = convert("input.pdf", output_md="out.md", flow="pymupdf4llm")
"""

import logging
import os
import re
import sys
import time
from dataclasses import dataclass

import pymupdf

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Optional imports — degrade gracefully
# ---------------------------------------------------------------------------
try:
    import cv2
    import numpy as np

    _HAS_CV2 = True
except ImportError:
    _HAS_CV2 = False

try:
    from doctr.io import DocumentFile
    from doctr.models import ocr_predictor

    _HAS_DOCTR = True
except ImportError:
    _HAS_DOCTR = False

try:
    from surya.detection import DetectionPredictor
    from surya.foundation import FoundationPredictor
    from surya.recognition import RecognitionPredictor

    _HAS_SURYA = True
except ImportError:
    _HAS_SURYA = False

try:
    import pymupdf4llm

    _HAS_PYMUPDF4LLM = True
except ImportError:
    _HAS_PYMUPDF4LLM = False

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

# Formatting types
FMT_NORMAL = 0
FMT_STRIKEOUT = 1  # → <del>
FMT_UNDERLINE = 2  # → <ins>

# char_flags bitmask positions (PyMuPDF >= 1.25.2)
CF_STRIKEOUT = 1  # bit 0
CF_UNDERLINE = 2  # bit 1

# Line-art classification tolerances (PDF points)
LINE_HEIGHT_MAX = 3.0  # max height of a line/rect to be considered formatting
MIN_LINE_WIDTH = 10.0  # ignore tiny line fragments

# Scanned page detection threshold
MIN_SELECTABLE_TEXT_LEN = 20

# ══════════════════════════════════════════════════════════════════════════════
# DATA STRUCTURE
# ══════════════════════════════════════════════════════════════════════════════


@dataclass
class FormattedSpan:
    text: str
    fmt: int  # FMT_NORMAL, FMT_STRIKEOUT, or FMT_UNDERLINE


# ══════════════════════════════════════════════════════════════════════════════
# LAYER 1: char_flags DETECTION
# ══════════════════════════════════════════════════════════════════════════════


def _detect_char_flags(span: dict) -> int:
    """Check span's char_flags for strikeout/underline."""
    cf = span.get("char_flags", 0)
    if cf & CF_STRIKEOUT:
        return FMT_STRIKEOUT
    if cf & CF_UNDERLINE:
        return FMT_UNDERLINE
    return FMT_NORMAL


# ══════════════════════════════════════════════════════════════════════════════
# LAYER 2: ANNOTATION DETECTION
# ══════════════════════════════════════════════════════════════════════════════


def _get_annotation_rects(page) -> tuple[list, list]:
    """Return (strikeout_rects, underline_rects) from page annotations."""
    strike_rects = []
    underline_rects = []
    try:
        for annot in page.annots():
            if annot.type[0] == pymupdf.PDF_ANNOT_STRIKE_OUT:
                strike_rects.append(annot.rect)
            elif annot.type[0] == pymupdf.PDF_ANNOT_UNDERLINE:
                underline_rects.append(annot.rect)
    except Exception:
        pass  # page may have no annots or broken annot data
    return strike_rects, underline_rects


def _classify_by_annotations(
    span_rect: pymupdf.Rect,
    strike_rects: list,
    underline_rects: list,
) -> int:
    """Check if span bbox intersects any annotation rect."""
    for sr in strike_rects:
        if span_rect.intersects(sr):
            return FMT_STRIKEOUT
    for ur in underline_rects:
        if span_rect.intersects(ur):
            return FMT_UNDERLINE
    return FMT_NORMAL


# ══════════════════════════════════════════════════════════════════════════════
# LAYER 3: LINE-ART / VECTOR GRAPHICS DETECTION
# ══════════════════════════════════════════════════════════════════════════════


@dataclass
class _FormattingLine:
    """A horizontal line detected in the page's vector graphics."""

    rect: pymupdf.Rect  # bounding rect of the line
    y_center: float  # vertical center of the line


def _get_formatting_lines(page) -> list[_FormattingLine]:
    """Extract horizontal lines/thin rects from page drawings."""
    lines = []
    try:
        drawings = page.get_drawings()
    except Exception:
        return lines

    for d in drawings:
        drect = pymupdf.Rect(d["rect"])
        # Must be thin and wide enough to be formatting
        if drect.height > LINE_HEIGHT_MAX:
            continue
        if drect.width < MIN_LINE_WIDTH:
            continue
        lines.append(
            _FormattingLine(rect=drect, y_center=(drect.y0 + drect.y1) / 2)
        )
    return lines


def _classify_by_drawings(
    span_rect: pymupdf.Rect,
    formatting_lines: list[_FormattingLine],
) -> int:
    """Check if any horizontal line overlaps the span and classify it."""
    if not formatting_lines:
        return FMT_NORMAL

    span_mid_y = (span_rect.y0 + span_rect.y1) / 2
    span_bottom_y = span_rect.y1
    span_height = span_rect.y1 - span_rect.y0
    # Tolerance relative to text height
    tol = max(span_height * 0.25, 2.0)

    for fl in formatting_lines:
        # Check horizontal overlap
        if fl.rect.x1 < span_rect.x0 or fl.rect.x0 > span_rect.x1:
            continue
        # Check if line overlaps span's vertical extent
        if fl.y_center < span_rect.y0 - tol or fl.y_center > span_rect.y1 + tol:
            continue

        # Classify by vertical position within the text bbox
        if abs(fl.y_center - span_mid_y) < tol:
            return FMT_STRIKEOUT
        if abs(fl.y_center - span_bottom_y) < tol:
            return FMT_UNDERLINE

    return FMT_NORMAL


# ══════════════════════════════════════════════════════════════════════════════
# DIGITAL PAGE EXTRACTION (Layers 1-3)
# ══════════════════════════════════════════════════════════════════════════════


def _extract_digital_page(
    page, exclude_rects: list[pymupdf.Rect] | None = None,
) -> list[list[FormattedSpan]]:
    """Extract text from a digital PDF page with formatting detection.

    Parameters
    ----------
    page          : PyMuPDF Page
    exclude_rects : regions to skip (e.g. table bounding boxes)

    Returns list of paragraphs, each a list of FormattedSpan.
    """
    data = page.get_text("dict")
    strike_rects, underline_rects = _get_annotation_rects(page)
    formatting_lines = _get_formatting_lines(page)
    exclude_rects = exclude_rects or []

    paragraphs: list[list[FormattedSpan]] = []
    for block in data["blocks"]:
        if block.get("type", 0) != 0:  # skip image blocks
            continue

        # Skip blocks that fall inside excluded regions (e.g. tables)
        block_rect = pymupdf.Rect(block["bbox"])
        if any(er.contains(block_rect) for er in exclude_rects):
            continue

        para_spans: list[FormattedSpan] = []
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "")
                if not text:
                    continue

                span_rect = pymupdf.Rect(span["bbox"])

                # Skip spans inside excluded regions
                if any(er.contains(span_rect) for er in exclude_rects):
                    continue

                # Layer 1: char_flags (most reliable)
                fmt = _detect_char_flags(span)

                # Layer 2: annotations
                if fmt == FMT_NORMAL and (strike_rects or underline_rects):
                    fmt = _classify_by_annotations(
                        span_rect, strike_rects, underline_rects
                    )

                # Layer 3: line-art
                if fmt == FMT_NORMAL and formatting_lines:
                    fmt = _classify_by_drawings(span_rect, formatting_lines)

                para_spans.append(FormattedSpan(text, fmt))

        if para_spans:
            paragraphs.append(para_spans)
    return paragraphs


# ══════════════════════════════════════════════════════════════════════════════
# LAYER 4: OCR + OpenCV FOR SCANNED PAGES
# ══════════════════════════════════════════════════════════════════════════════


def _is_scanned_page(page) -> bool:
    """Check if page lacks selectable text (likely scanned/image-based)."""
    return len(page.get_text("text").strip()) < MIN_SELECTABLE_TEXT_LEN


# --- docTR OCR ---------------------------------------------------------------

# Lazy-init singleton so model is loaded once across all pages
_doctr_model = None


def _get_doctr_model():
    global _doctr_model
    if _doctr_model is None:
        log.info("Loading docTR model (first call)...")
        _doctr_model = ocr_predictor(
            det_arch="db_resnet50",
            reco_arch="crnn_vgg16_bn",
            pretrained=True,
            assume_straight_pages=True,
        )
    return _doctr_model


def _ocr_page_doctr(page, dpi: int = 300) -> list[dict]:
    """OCR a page using docTR. Returns [{"text": str, "bbox": Rect}, ...]."""
    pix = page.get_pixmap(dpi=dpi)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    if pix.n == 4:  # RGBA → RGB
        img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
    elif pix.n == 1:  # Gray → RGB
        img = cv2.cvtColor(img, cv2.COLOR_GRAY2RGB)

    model = _get_doctr_model()
    # docTR 1.x from_images expects file paths or bytes, not numpy arrays
    _, buf = cv2.imencode(".png", img)
    doc = DocumentFile.from_images([buf.tobytes()])
    result = model(doc)

    scale_x = page.rect.width
    scale_y = page.rect.height
    words = []
    for p in result.pages:
        for block in p.blocks:
            for line in block.lines:
                for word in line.words:
                    # docTR returns normalized coords (0-1)
                    (x0, y0), (x1, y1) = word.geometry
                    bbox = pymupdf.Rect(
                        x0 * scale_x,
                        y0 * scale_y,
                        x1 * scale_x,
                        y1 * scale_y,
                    )
                    words.append({"text": word.value, "bbox": bbox})
    return words


# --- Tesseract OCR (via PyMuPDF built-in) ------------------------------------


def _ocr_page_tesseract(page, dpi: int = 300) -> list[dict]:
    """OCR a page using PyMuPDF's built-in Tesseract. Returns word dicts."""
    tp = page.get_textpage_ocr(dpi=dpi, language="eng", full=True)
    raw_words = page.get_text("words", textpage=tp)
    words = []
    for w in raw_words:
        bbox = pymupdf.Rect(w[:4])
        text = w[4]
        if text.strip():
            words.append({"text": text, "bbox": bbox})
    return words


# --- Surya OCR ----------------------------------------------------------------

# Lazy-init singletons so models are loaded once across all pages
_surya_det = None
_surya_rec = None


def _get_surya_models():
    global _surya_det, _surya_rec
    if _surya_rec is None:
        log.info("Loading Surya models (first call)...")
        foundation = FoundationPredictor()
        _surya_det = DetectionPredictor()
        _surya_rec = RecognitionPredictor(foundation)
    return _surya_det, _surya_rec


def _ocr_page_surya(page, dpi: int = 300) -> list[dict]:
    """OCR a page using Surya. Returns [{"text": str, "bbox": Rect}, ...]."""
    from PIL import Image

    pix = page.get_pixmap(dpi=dpi)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    if pix.n == 4:  # RGBA → RGB
        img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
    elif pix.n == 1:  # Gray → RGB
        img = cv2.cvtColor(img, cv2.COLOR_GRAY2RGB)

    pil_img = Image.fromarray(img)
    det, rec = _get_surya_models()
    results = rec([pil_img], det_predictor=det, return_words=True, math_mode=False)

    # Surya returns pixel coords at the rendered DPI — scale to PDF points
    scale = 72.0 / dpi
    words = []
    for line in results[0].text_lines:
        for w in (line.words or []):
            if not w.bbox_valid:
                continue
            x0, y0, x1, y1 = w.bbox
            bbox = pymupdf.Rect(
                x0 * scale, y0 * scale, x1 * scale, y1 * scale
            )
            words.append({"text": w.text, "bbox": bbox})
    return words


# --- OpenCV horizontal line detection ----------------------------------------


def _detect_lines_opencv(page, dpi: int = 300) -> list[_FormattingLine]:
    """Render page to image, detect horizontal lines via morphological ops.

    Returns list of _FormattingLine in PDF-point coordinates.
    """
    if not _HAS_CV2:
        log.warning("OpenCV not installed — skipping line detection for scanned page")
        return []

    pix = page.get_pixmap(dpi=dpi)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    if pix.n >= 3:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    else:
        gray = img

    # Binary threshold — dark lines on light background
    _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

    # Morphological open with wide horizontal kernel to isolate horizontal lines
    kernel_width = max(40, pix.w // 30)
    horizontal_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT, (kernel_width, 1)
    )
    lines_mask = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel)

    # Find contours
    contours, _ = cv2.findContours(
        lines_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )

    scale = dpi / 72.0  # pixel → PDF-point factor
    fmt_lines = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w / scale < MIN_LINE_WIDTH:
            continue
        # Convert pixel coords to PDF points
        rect = pymupdf.Rect(x / scale, y / scale, (x + w) / scale, (y + h) / scale)
        fmt_lines.append(
            _FormattingLine(rect=rect, y_center=(rect.y0 + rect.y1) / 2)
        )
    return fmt_lines


# --- Scanned page assembly ---------------------------------------------------


def _group_words_into_paragraphs(
    words: list[dict], line_gap_factor: float = 1.5
) -> list[list[dict]]:
    """Group OCR'd words into paragraphs by Y proximity.

    Words on the same line (similar Y) form a line.
    Lines separated by more than line_gap_factor * avg_line_height start a new paragraph.
    """
    if not words:
        return []

    # Sort by Y then X
    sorted_words = sorted(words, key=lambda w: (w["bbox"].y0, w["bbox"].x0))

    # Group into lines (words with overlapping Y ranges)
    lines: list[list[dict]] = []
    current_line = [sorted_words[0]]
    for w in sorted_words[1:]:
        prev = current_line[-1]
        # Same line if Y overlaps significantly
        prev_mid = (prev["bbox"].y0 + prev["bbox"].y1) / 2
        curr_mid = (w["bbox"].y0 + w["bbox"].y1) / 2
        avg_h = (
            prev["bbox"].y1 - prev["bbox"].y0 + w["bbox"].y1 - w["bbox"].y0
        ) / 2
        if abs(curr_mid - prev_mid) < avg_h * 0.6:
            current_line.append(w)
        else:
            # Sort line words by X
            current_line.sort(key=lambda _w: _w["bbox"].x0)
            lines.append(current_line)
            current_line = [w]
    if current_line:
        current_line.sort(key=lambda _w: _w["bbox"].x0)
        lines.append(current_line)

    # Group lines into paragraphs
    if not lines:
        return []

    paragraphs: list[list[dict]] = []
    current_para_lines = [lines[0]]
    for i in range(1, len(lines)):
        prev_line = current_para_lines[-1]
        curr_line = lines[i]
        prev_bottom = max(w["bbox"].y1 for w in prev_line)
        curr_top = min(w["bbox"].y0 for w in curr_line)
        avg_h = sum(
            w["bbox"].y1 - w["bbox"].y0 for w in prev_line
        ) / len(prev_line)
        gap = curr_top - prev_bottom
        if gap > avg_h * line_gap_factor:
            # Flatten lines into one paragraph word list
            paragraphs.append([w for ln in current_para_lines for w in ln])
            current_para_lines = [curr_line]
        else:
            current_para_lines.append(curr_line)
    if current_para_lines:
        paragraphs.append([w for ln in current_para_lines for w in ln])

    return paragraphs


def _classify_word_by_lines(
    word_bbox: pymupdf.Rect,
    formatting_lines: list[_FormattingLine],
) -> int:
    """Classify a word as strikethrough/underline based on overlapping lines."""
    word_mid_y = (word_bbox.y0 + word_bbox.y1) / 2
    word_bottom_y = word_bbox.y1
    word_height = word_bbox.y1 - word_bbox.y0
    tol = max(word_height * 0.25, 2.0)

    for fl in formatting_lines:
        # Horizontal overlap
        if fl.rect.x1 < word_bbox.x0 or fl.rect.x0 > word_bbox.x1:
            continue
        # Vertical proximity
        if fl.y_center < word_bbox.y0 - tol or fl.y_center > word_bbox.y1 + tol:
            continue
        if abs(fl.y_center - word_mid_y) < tol:
            return FMT_STRIKEOUT
        if abs(fl.y_center - word_bottom_y) < tol:
            return FMT_UNDERLINE
    return FMT_NORMAL


def _extract_scanned_page(
    page, dpi: int = 300, ocr_engine: str = "doctr"
) -> list[list[FormattedSpan]]:
    """Extract text from a scanned page using OCR + OpenCV line detection."""
    # OCR
    if ocr_engine == "surya" and _HAS_SURYA and _HAS_CV2:
        words = _ocr_page_surya(page, dpi)
    elif ocr_engine == "surya" and not _HAS_SURYA:
        log.warning("Surya not installed — falling back to Tesseract")
        words = _ocr_page_tesseract(page, dpi)
    elif ocr_engine == "doctr" and _HAS_DOCTR and _HAS_CV2:
        words = _ocr_page_doctr(page, dpi)
    elif ocr_engine == "doctr" and not _HAS_DOCTR:
        log.warning("docTR not installed — falling back to Tesseract")
        words = _ocr_page_tesseract(page, dpi)
    else:
        words = _ocr_page_tesseract(page, dpi)

    if not words:
        return []

    # Detect horizontal lines
    formatting_lines = _detect_lines_opencv(page, dpi)

    # Group words into paragraphs
    para_words = _group_words_into_paragraphs(words)

    paragraphs: list[list[FormattedSpan]] = []
    for pw in para_words:
        spans = []
        for w in pw:
            fmt = _classify_word_by_lines(w["bbox"], formatting_lines)
            spans.append(FormattedSpan(w["text"], fmt))
        if spans:
            paragraphs.append(spans)
    return paragraphs


# ══════════════════════════════════════════════════════════════════════════════
# TABLE DETECTION
# ══════════════════════════════════════════════════════════════════════════════


def _extract_tables_as_markdown(page) -> tuple[str, list[pymupdf.Rect]]:
    """Detect tables via PyMuPDF, render as markdown tables with formatting.

    Returns (markdown_string, list_of_table_rects) so we can exclude those
    regions from paragraph extraction.
    """
    try:
        tables = page.find_tables()
    except Exception:
        return "", []

    if not tables or not tables.tables:
        return "", []

    # Pre-compute formatting info for classification inside table cells
    strike_rects, underline_rects = _get_annotation_rects(page)
    formatting_lines = _get_formatting_lines(page)

    md_parts = []
    table_rects = []
    for table in tables.tables:
        table_rects.append(pymupdf.Rect(table.bbox))
        extracted = table.extract()
        if not extracted:
            continue
        # Build markdown table
        header = extracted[0]
        col_count = len(header)
        # Header row
        header_cells = []
        for cell in header:
            header_cells.append(cell if cell else "")
        md_parts.append("| " + " | ".join(header_cells) + " |")
        md_parts.append("| " + " | ".join(["---"] * col_count) + " |")
        # Data rows
        for row in extracted[1:]:
            cells = []
            for cell in row:
                cells.append(cell if cell else "")
            md_parts.append("| " + " | ".join(cells) + " |")
        md_parts.append("")  # blank line after table

    return "\n".join(md_parts), table_rects



# ══════════════════════════════════════════════════════════════════════════════
# MARKDOWN ASSEMBLY
# ══════════════════════════════════════════════════════════════════════════════


def _spans_to_markdown(
    paragraphs: list[list[FormattedSpan]], word_sep: str = "",
) -> str:
    """Convert formatted spans into markdown with <ins>/<del> tags.

    Groups consecutive same-format spans into one tag pair.

    Parameters
    ----------
    paragraphs : list of paragraphs, each a list of FormattedSpan
    word_sep   : separator between spans. Use " " for OCR'd words, "" for
                 digital spans (PyMuPDF spans already include trailing spaces).
    """
    md_parts = []
    for para in paragraphs:
        if not para:
            continue
        line_parts = []
        i = 0
        while i < len(para):
            span = para[i]
            fmt = span.fmt

            if fmt == FMT_NORMAL:
                line_parts.append(span.text)
                i += 1
            else:
                # Collect consecutive spans with the same formatting
                group_texts = [span.text]
                j = i + 1
                while j < len(para) and para[j].fmt == fmt:
                    group_texts.append(para[j].text)
                    j += 1
                merged = word_sep.join(group_texts) if word_sep else "".join(group_texts)
                if fmt == FMT_STRIKEOUT:
                    line_parts.append(f"<del>{merged}</del>")
                elif fmt == FMT_UNDERLINE:
                    line_parts.append(f"<ins>{merged}</ins>")
                i = j

        para_text = word_sep.join(line_parts).strip() if word_sep else "".join(line_parts).strip()
        if para_text:
            md_parts.append(para_text)

    return "\n\n".join(md_parts)


# ══════════════════════════════════════════════════════════════════════════════
# CLEANUP
# ══════════════════════════════════════════════════════════════════════════════


def clean_markdown(md: str) -> str:
    """Merge adjacent same-type tags and remove empty ones."""
    # Merge adjacent tags: </ins>whitespace<ins> → keep whitespace
    md = re.sub(r"</ins>(\s*)<ins>", r"\1", md)
    md = re.sub(r"</del>(\s*)<del>", r"\1", md)

    # Remove empty tags
    md = re.sub(r"<ins>\s*</ins>", "", md)
    md = re.sub(r"<del>\s*</del>", "", md)

    # Clean up extra blank lines
    md = re.sub(r"\n{3,}", "\n\n", md)

    return md.strip()


# ══════════════════════════════════════════════════════════════════════════════
# PYMUPDF4LLM FLOW
# ══════════════════════════════════════════════════════════════════════════════


def _convert_pymupdf4llm(input_pdf: str, output_md: str | None = None) -> str:
    """Force-OCR the entire PDF via pymupdf4llm and return raw markdown.

    No <ins>/<del> detection — use this flow when you need clean markdown
    from scanned documents without track-change annotation.
    """
    if not _HAS_PYMUPDF4LLM:
        raise ImportError(
            "pymupdf4llm not installed — run: pip install pymupdf4llm"
        )

    t0 = time.time()
    md = pymupdf4llm.to_markdown(input_pdf, force_ocr=True)
    elapsed = time.time() - t0
    log.info("pymupdf4llm flow: %.1fs for %s", elapsed, input_pdf)

    if output_md:
        with open(output_md, "w", encoding="utf-8") as f:
            f.write(md)

    return md


# ══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ══════════════════════════════════════════════════════════════════════════════


def convert(
    input_pdf: str,
    output_md: str = None,
    ocr_engine: str = "surya",
    dpi: int = 300,
    flow: str = "default",
) -> str:
    """
    Convert PDF with underline/strikethrough to Markdown with <ins>/<del> tags.

    Parameters
    ----------
    input_pdf  : path to the PDF file
    output_md  : path to write the .md file (optional)
    ocr_engine : "surya" (default), "doctr", or "tesseract"
    dpi        : DPI for OCR / image rendering (default 300)
    flow       : "default" (layered ins/del detection) or "pymupdf4llm"
                 (force-OCR whole document via pymupdf4llm, no ins/del tags)

    Returns
    -------
    Markdown string. "default" flow includes <ins>/<del> HTML tags.
    "pymupdf4llm" flow returns raw markdown without track-change markup.
    """
    if flow == "pymupdf4llm":
        return _convert_pymupdf4llm(input_pdf, output_md)

    from tqdm import tqdm

    if not os.path.isfile(input_pdf):
        raise FileNotFoundError(f"PDF not found: {input_pdf}")

    doc = pymupdf.open(input_pdf)
    total_pages = len(doc)
    all_md_parts: list[str] = []

    t0 = time.time()
    scanned_count = 0
    digital_count = 0

    for page in tqdm(doc, desc="Processing pages", total=total_pages):
        page_md_parts = []

        # Detect tables first (so we can exclude those regions)
        table_md, table_rects = _extract_tables_as_markdown(page)

        if _is_scanned_page(page):
            scanned_count += 1
            paragraphs = _extract_scanned_page(page, dpi, ocr_engine)
            word_sep = " "  # OCR'd words need explicit spaces
        else:
            digital_count += 1
            paragraphs = _extract_digital_page(page, exclude_rects=table_rects)
            word_sep = ""  # PyMuPDF spans include trailing spaces

        para_md = _spans_to_markdown(paragraphs, word_sep=word_sep)
        if para_md:
            page_md_parts.append(para_md)
        if table_md:
            page_md_parts.append(table_md)

        if page_md_parts:
            all_md_parts.append("\n\n".join(page_md_parts))

    doc.close()

    md = "\n\n---\n\n".join(all_md_parts)  # page separator
    md = clean_markdown(md)

    elapsed = time.time() - t0
    log.info(
        "Run summary: %d pages (digital=%d, scanned=%d) in %.1fs",
        total_pages,
        digital_count,
        scanned_count,
        elapsed,
    )

    if output_md:
        with open(output_md, "w", encoding="utf-8") as f:
            f.write(md)

    return md


# ══════════════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    parser = argparse.ArgumentParser(
        description="Convert PDF with underline/strikethrough to Markdown with <ins>/<del> tags."
    )
    parser.add_argument("input", help="Path to the PDF file")
    parser.add_argument(
        "output",
        nargs="?",
        default=None,
        help="Path to write the .md file (default: input with .md extension)",
    )
    parser.add_argument(
        "--ocr-engine",
        choices=["surya", "doctr", "tesseract"],
        default="surya",
        help="OCR engine for scanned pages (default: surya)",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="DPI for OCR / image rendering (default: 300)",
    )
    parser.add_argument(
        "--flow",
        choices=["default", "pymupdf4llm"],
        default="default",
        help="Conversion flow: 'default' (layered ins/del detection) or "
             "'pymupdf4llm' (force-OCR via pymupdf4llm, no ins/del tags)",
    )
    args = parser.parse_args()

    src = args.input
    dst = args.output if args.output else src.rsplit(".", 1)[0] + ".md"

    result = convert(src, dst, ocr_engine=args.ocr_engine, dpi=args.dpi, flow=args.flow)
    print(result)
