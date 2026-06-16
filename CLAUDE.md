# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

Converts document files with track changes into Markdown with `<ins>`/`<del>` HTML tags preserving the change markup. Supports DOCX, DOC, and PDF inputs.

### DOCX/DOC Pipeline (`docx_parsing.py`)

Converts `.docx` (or `.doc`) files with Word track changes (insertions, deletions, moves, formatting changes). The pipeline:

0. If input is `.doc`:
   - **Extract WordML equations** from the OLE `Data` stream (preserves inline math as LaTeX — see README for details)
   - Convert `.doc` to `.docx` via LibreOffice headless mode
1. Unzips the `.docx`, parses the OOXML, and injects Unicode marker strings around tracked changes
2. Reassembles a modified `.docx` and converts it to Markdown via pandoc
3. Splices extracted LaTeX into image placeholders (`.doc` equations only) and replaces markers with HTML tags

### PDF Pipeline (`pdf_parsing.py`)

Converts PDFs where visual strikethrough = deletion and underline = insertion (common in legal/standards documents). Uses a 4-layer detection strategy:

1. **char_flags** — PyMuPDF MuPDF-level detection (bit 0=strikeout, bit 1=underline)
2. **Annotations** — PDF annotation objects (`PDF_ANNOT_STRIKE_OUT`, `PDF_ANNOT_UNDERLINE`)
3. **Line-art** — Horizontal lines/thin rects from page vector graphics (`get_drawings()`)
4. **OCR+OpenCV** — For scanned pages: Surya/docTR/Tesseract OCR + morphological line detection

## Dependencies

Install all Python dependencies: `pip install -r requirements.txt`

### Core (always required)
- **Python packages:** `tqdm`, `openpyxl` (batch mode Excel summaries)

### DOCX/DOC
- **Python packages:** `lxml`, `pypandoc`, `olefile` (`.doc` equation extraction)
- **System dependencies:** `pandoc` (required), `libreoffice` (only for `.doc` input)
- **Optional — equation OCR:** `pix2tex` (`pip install pix2tex`) — fallback for `--ocr-equations`. Only useful if the `.doc` has no WordML math embedded. For most modern `.doc` files the built-in WordML extractor recovers equations as LaTeX automatically.

### PDF
- **Python packages:** `pymupdf`, `opencv-python-headless`, `surya-ocr`, `transformers>=4.56,<5`
- **Alternative OCR engines:** `python-doctr[torch]` (pip), `tesseract` (system install)

## Directory Layout

```
.
├── input/      ← place source files here (.doc, .docx, .pdf)
├── output/     ← converted Markdown written here by default
│   ├── .convert_state.json          ← persistent skip-tracking (batch mode)
│   └── conversion_summary_*.xlsx    ← per-run Excel reports (batch mode)
├── convert.py
├── docx_parsing.py
├── pdf_parsing.py
└── ...
```

## Running

```bash
# ── Single-file mode ────────────────────────────────────────────────
# Auto-detects file type; output goes to output/<stem>.md
python convert.py input/file.docx
python convert.py input/file.doc
python convert.py input/file.doc --ocr-equations        # OCR equations via pix2tex
python convert.py input/file.pdf
python convert.py input/file.pdf --ocr-engine surya
python convert.py input/file.pdf --ocr-engine doctr
python convert.py input/file.pdf --ocr-engine tesseract

# Explicit output path
python convert.py input/file.docx output/custom.md

# ── Batch / folder mode ─────────────────────────────────────────────
# Converts all .doc/.docx/.pdf in a directory; skips already-converted files
python convert.py input/folder/
python convert.py input/folder/ output/folder/
python convert.py input/folder/ --ocr-engine tesseract
python convert.py input/folder/ --ocr-equations         # OCR equations in all .doc files

# ── Direct pipeline scripts ─────────────────────────────────────────
python docx_parsing.py input/file.docx [output/file.md]
python docx_parsing.py input/file.doc  [output/file.md] --ocr-equations
python pdf_parsing.py input/file.pdf [output/file.md]

# ── As a library ────────────────────────────────────────────────────
from convert import convert, batch_convert
md    = convert("input/file.pdf", ocr_engine="surya")
md    = convert("input/file.pdf", output_md="output/custom.md", ocr_engine="surya")
md    = convert("input/file.doc", ocr_equations=True)   # equation OCR for .doc
excel = batch_convert("input/folder/")
excel = batch_convert("input/folder/", output_dir="output/folder/", ocr_engine="tesseract")
excel = batch_convert("input/folder/", ocr_equations=True)
```

## Architecture Notes

### convert.py

Unified entry point. When the CLI argument is a **directory**, delegates to `batch_convert()`; when it is a **file**, delegates to `convert()`. Imports are deferred so only the needed pipeline's dependencies are loaded.

#### Batch mode (`batch_convert`)

- Scans input directory for `.doc`, `.docx`, `.pdf` files (sorted alphabetically).
- Loads `<output_dir>/.convert_state.json` to determine which files have already been converted. Skip key is `"<filename>:<size_bytes>"` — the size component allows two files with the same name but different content to be treated as distinct.
- Processes only new files; already-converted files still appear in the Excel report (marked "skipped (already converted)").
- State is saved **after each file** so a mid-run interruption doesn't lose progress.
- Output stem collisions (two input files sharing the same stem, e.g. `report.docx` and `report.pdf`) are resolved by appending `_1`, `_2`, etc.
- After all files, writes `<output_dir>/conversion_summary_YYYYMMDD_HHMMSS.xlsx` with:
  - **File Details** sheet — one row per file with input/output paths, size, method, OCR engine, status, conversion datetime, elapsed seconds, error.
  - **Run Summary** sheet — aggregate counts and run metadata.

### docx_parsing.py

Single-file tool. Key design decisions:

- **`.doc` → `.docx` conversion:** Uses LibreOffice headless (`soffice --headless --convert-to docx`) with a temporary user profile dir to avoid lock conflicts. `_find_soffice()` auto-detects the binary at the macOS app path (`/Applications/LibreOffice.app/Contents/MacOS/soffice`), Linux path (`/usr/bin/soffice`), or anywhere on `PATH`. Adapted from the server batch script `document_conversion_doc_to_docx.py` — stripped down to single-file conversion without multiprocessing, config.json, or process killing. Works on both macOS (local dev) and Linux (production server).
- **WordML equation extraction from `.doc` (primary equation path):** Many `.doc` files saved by modern Word versions embed Word-2003 WordML XML snippets in the OLE `Data` stream — one `<m:oMath>` element per inline equation. LibreOffice **discards these entirely** when converting to `.docx` (equations become transparent 0-pixel PNGs). Before the LibreOffice step, `extract_doc_wordml_equations()` opens the `.doc` via `olefile`, scans the `Data` stream for embedded `<?xml ...>…</w:wordDocument>` blobs, dedupes them (Word writes each object twice), unwraps any `<aml:annotation>` tracked-change wrappers inside the math (otherwise pandoc drops the content), packs each `<m:oMath>` into a minimal in-memory `.docx`, and runs pandoc to get LaTeX. After pandoc converts the full document, `_replace_images_with_latex()` splices the N-th extracted LaTeX string into the N-th `![](media/imageN.png)` placeholder in document order. This path produces real `$LaTeX$` output without any OCR.
- **Equation OCR (`--ocr-equations`, fallback):** Used only when the `.doc` has no WordML math embedded (older Equation Editor OLE or true MathType-binary `.doc`). When `ocr_equations=True`, `_extract_media()` reads the equation PNGs into memory before the temp dir is torn down, then `_replace_equation_images()` post-processes the pandoc markdown: any `![](media/imageN.png)` with `height < 0.5in` is passed to pix2tex and replaced with `$<latex>$`. The pix2tex model is lazily loaded and cached module-globally so it is only initialised once per process. Note: OCR cannot help when LibreOffice emits blank transparent PNGs (pixel values all zero) — that case requires the WordML extractor to succeed.
- **Marker strategy:** Unicode math brackets (`⟦TRACK_INS⟧` etc.) are injected into the OOXML as `<w:t>` text runs *before* pandoc conversion. These survive pandoc intact and are string-replaced with HTML tags afterward. This avoids fighting pandoc's XML handling.
- **Processing order matters:** Moves are processed first (they contain nested ins/del), then insertions, then deletions, then format changes. `process_change_type` uses a `while True` loop re-querying XPath each iteration because injecting markers mutates the tree and invalidates prior references.
- **Block vs. inline injection:** Changes inside `<w:p>` (paragraphs, hyperlinks, etc.) get inline marker runs inserted. Changes at body/table-cell level get markers placed inside the first/last `<w:p>` child, then children are unwrapped to the parent.
- **`_XML_PART_RE`** controls which `.docx` internal files are processed — document body, headers, footers, footnotes, endnotes.

### pdf_parsing.py

Single-file tool. Key design decisions:

- **Layered detection:** char_flags (Layer 1) → annotations (Layer 2) → line-art (Layer 3) → OCR+OpenCV (Layer 4). Each layer is tried in order per span; first match wins.
- **Direct PyMuPDF extraction** rather than pymupdf4llm wrapper: gives authoritative per-span `char_flags` without fragile text-matching. pymupdf4llm doesn't support underline in markdown output.
- **char_flags** (bit 0=strikeout, bit 1=underline) available since PyMuPDF 1.25.2 via `page.get_text("dict")` — most reliable for digital PDFs.
- **Line-art classification:** Horizontal lines detected via `page.get_drawings()` are classified as strikethrough (y near text midpoint) or underline (y near text bottom) using relative tolerance.
- **OCR engines:** Three engines supported — Surya (default, best text accuracy, GPL-3.0), docTR (good formatting detection, Apache-2.0), Tesseract (fastest, via PyMuPDF built-in). Surya requires `transformers>=4.56,<5` (not 5.x). docTR 1.x `DocumentFile.from_images()` requires PNG bytes, not raw numpy arrays.
- **Scanned page detection:** Pages with <20 chars of selectable text are treated as scanned.
- **OpenCV line detection:** Morphological open with wide horizontal kernel isolates horizontal lines in page images. Lines are matched to OCR'd word bounding boxes and classified by Y position.
- **Table handling:** `page.find_tables()` detects tables; their bounding rects are excluded from paragraph extraction to avoid duplication.

## Mistakes Log
