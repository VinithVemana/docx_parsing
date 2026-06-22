"""
docx_track_changes_to_md.py

Convert .docx (or .doc) with track changes → Markdown with <ins>/<del> tags.

Requirements:
    pip install lxml pypandoc
    pandoc must be installed: https://pandoc.org/installing.html
    LibreOffice must be installed for .doc support: brew install --cask libreoffice

Usage:
    python docx_parsing.py input.docx [output.md]
    python docx_parsing.py input.doc  [output.md]
    python docx_parsing.py input.doc  [output.md] --ocr-equations   # OCR equations via pix2tex
    python docx_parsing.py input.docx [output.md] --no-heading-fix  # disable RAG heading post-process
    python docx_parsing.py input.docx [output.md] -f gfm            # pandoc format other than markdown
"""

import io
import os
import re
import shutil
import subprocess
import sys
import zipfile
import tempfile
from lxml import etree
import pypandoc

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS  = "http://www.w3.org/XML/1998/namespace"
NSMAP   = {"w": WORD_NS}

def qn(tag):
    """Qualify a tag name with the Word namespace."""
    return f"{{{WORD_NS}}}{tag}"

# Unique markers — Unicode math brackets won't appear in normal text
# and survive pandoc conversion intact
MARKERS = {
    "ins":       ("⟦TRACK_INS⟧",       "⟦/TRACK_INS⟧"),
    "del":       ("⟦TRACK_DEL⟧",       "⟦/TRACK_DEL⟧"),
    "move_from": ("⟦TRACK_MOVEFROM⟧",  "⟦/TRACK_MOVEFROM⟧"),
    "move_to":   ("⟦TRACK_MOVETO⟧",    "⟦/TRACK_MOVETO⟧"),
    "fmt":       ("⟦TRACK_FMT⟧",       "⟦/TRACK_FMT⟧"),
}

# What each marker becomes in the final Markdown
TAG_MAP = {
    "⟦TRACK_INS⟧":       "<ins>",
    "⟦/TRACK_INS⟧":      "</ins>",
    "⟦TRACK_DEL⟧":       "<del>",
    "⟦/TRACK_DEL⟧":      "</del>",
    "⟦TRACK_MOVEFROM⟧":  '<del class="move-from">',
    "⟦/TRACK_MOVEFROM⟧": "</del>",
    "⟦TRACK_MOVETO⟧":    '<ins class="move-to">',
    "⟦/TRACK_MOVETO⟧":   "</ins>",
    "⟦TRACK_FMT⟧":       '<span class="fmt-change">',
    "⟦/TRACK_FMT⟧":      "</span>",
}


# ══════════════════════════════════════════════════════════════════════════════
# XML HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def make_text_run(text):
    """Create a <w:r><w:t xml:space='preserve'>text</w:t></w:r>."""
    r = etree.Element(qn("r"))
    t = etree.SubElement(r, qn("t"))
    t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text
    return r


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1: PROCESS TRACKED CHANGES IN THE XML
# ══════════════════════════════════════════════════════════════════════════════

def _inject_inline(elem, open_m, close_m, convert_del_text):
    """
    Inline change (parent is <w:p>, <w:hyperlink>, etc.).
    Replace the wrapper element with:  [open_marker] [children...] [close_marker]
    """
    parent = elem.getparent()
    idx = list(parent).index(elem)

    if convert_del_text:
        for dt in elem.iter(qn("delText")):
            dt.tag = qn("t")  # make deleted text visible

    children = list(elem)
    parent.remove(elem)

    parent.insert(idx, make_text_run(open_m))
    for i, child in enumerate(children):
        parent.insert(idx + 1 + i, child)
    parent.insert(idx + 1 + len(children), make_text_run(close_m))


def _inject_block(elem, open_m, close_m, convert_del_text):
    """
    Block-level change (parent is <w:body>, <w:tc>, etc.).
    Insert markers inside the first and last <w:p> found within,
    then unwrap children to parent.
    """
    parent = elem.getparent()
    idx = list(parent).index(elem)

    if convert_del_text:
        for dt in elem.iter(qn("delText")):
            dt.tag = qn("t")

    paras = list(elem.iter(qn("p")))
    if paras:
        paras[0].insert(0, make_text_run(open_m))
        paras[-1].append(make_text_run(close_m))

    children = list(elem)
    parent.remove(elem)
    for i, child in enumerate(children):
        parent.insert(idx + i, child)


# Tags whose children are inline (runs, etc.)
_INLINE_PARENTS = {"p", "hyperlink", "smartTag", "fldSimple", "sdtContent"}

def process_change_type(root, xpath, marker_key, convert_del_text=False):
    """Find all elements matching xpath and inject markers around their content."""
    open_m, close_m = MARKERS[marker_key]

    while True:
        found = root.xpath(xpath, namespaces=NSMAP)
        if not found:
            break

        elem = found[0]
        parent = elem.getparent()

        if parent is None:
            break

        parent_tag = etree.QName(parent.tag).localname

        if parent_tag in _INLINE_PARENTS:
            _inject_inline(elem, open_m, close_m, convert_del_text)
        else:
            _inject_block(elem, open_m, close_m, convert_del_text)


def process_format_changes(root):
    """
    Handle <w:rPrChange> — mark runs whose formatting changed.
    We wrap the parent run with format-change markers.
    """
    open_m, close_m = MARKERS["fmt"]

    for rpr_change in list(root.iter(qn("rPrChange"))):
        rpr = rpr_change.getparent()
        if rpr is None:
            continue
        run = rpr.getparent()
        if run is None or etree.QName(run.tag).localname != "r":
            continue
        parent = run.getparent()
        if parent is None:
            continue

        idx = list(parent).index(run)
        parent.insert(idx, make_text_run(open_m))
        # run is now at idx+1, so close marker goes at idx+2
        parent.insert(idx + 2, make_text_run(close_m))

        # Remove rPrChange so it doesn't confuse pandoc
        rpr.remove(rpr_change)


def remove_property_change_elements(root):
    """
    Remove structural *Change elements (paragraph, section, table properties).
    These represent formatting changes to containers — we accept the new format.
    """
    for tag_name in (
        "pPrChange", "sectPrChange",                   # paragraph / section
        "tblPrChange", "trPrChange", "tcPrChange",     # table / row / cell
        "tblGridChange",                                # table grid
    ):
        for elem in list(root.iter(qn(tag_name))):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)


def process_all_changes(xml_bytes):
    """
    Master function: process all tracked changes in a document XML part.
    Returns modified XML as bytes.
    """
    root = etree.fromstring(xml_bytes)

    # 1. Moves (before ins/del — moves contain nested ins/del markers)
    process_change_type(root, ".//w:moveFrom", "move_from", convert_del_text=True)
    process_change_type(root, ".//w:moveTo",   "move_to",   convert_del_text=False)

    # 2. Insertions
    process_change_type(root, ".//w:ins", "ins", convert_del_text=False)

    # 3. Deletions
    process_change_type(root, ".//w:del", "del", convert_del_text=True)

    # 4. Formatting changes on runs
    process_format_changes(root)

    # 5. Clean up structural property changes
    remove_property_change_elements(root)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2: BUILD A MODIFIED .DOCX WITH MARKERS BAKED IN
# ══════════════════════════════════════════════════════════════════════════════

# Which files inside the .docx to process
_XML_PART_RE = re.compile(
    r"^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$"
)

def create_marked_docx(src_path, dst_path):
    """Copy a .docx, injecting markers into every relevant XML part."""
    with zipfile.ZipFile(src_path, "r") as zin, \
         zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:

        for item in zin.namelist():
            data = zin.read(item)

            if _XML_PART_RE.match(item):
                try:
                    data = process_all_changes(data)
                except Exception as exc:
                    print(f"⚠ Could not process {item}: {exc}", file=sys.stderr)

            zout.writestr(item, data)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3: PANDOC CONVERSION  +  MARKER → HTML TAG REPLACEMENT
# ══════════════════════════════════════════════════════════════════════════════

def markers_to_tags(md: str) -> str:
    """Replace ⟦MARKER⟧ placeholders with HTML tags."""
    for marker, tag in TAG_MAP.items():
        md = md.replace(marker, tag)

        # Pandoc sometimes backslash-escapes Unicode brackets
        escaped = marker.replace("⟦", "\\⟦").replace("⟧", "\\⟧")
        md = md.replace(escaped, tag)

    return md


def clean_markdown(md: str) -> str:
    """Merge adjacent same-type tags and remove empty ones."""
    # </ins>  <ins> → (merge into one span)
    md = re.sub(r"</ins>(\s*)<ins>",                   r"\1", md)
    md = re.sub(r"</del>(\s*)<del>",                   r"\1", md)
    md = re.sub(r'</del>(\s*)<del class="move-from">', r"\1", md)
    md = re.sub(r'</ins>(\s*)<ins class="move-to">',   r"\1", md)

    # Remove empty tags
    md = re.sub(r"<ins[^>]*>\s*</ins>", "", md)
    md = re.sub(r"<del[^>]*>\s*</del>", "", md)

    return md


# ══════════════════════════════════════════════════════════════════════════════
# HEADING POST-PROCESSING  (for RAG-friendly section segmentation)
# ══════════════════════════════════════════════════════════════════════════════

# False-positive headings: a single '#' immediately followed by a digit, e.g.
# '#3 (Illegal UE);' from CT1 reject-cause lists. Pandoc's permissive ATX
# parser treats these as H1, which over-segments RAG chunks.
_FP_HEAD_RE = re.compile(r"^(#+)(\d)")

# A real ATX heading: 1-6 '#' followed by whitespace and non-whitespace.
_REAL_HEAD_RE = re.compile(r"^(#{1,6})\s+\S")

# An empty ATX heading line: '#', '##', '#### ' with whitespace but no title.
# Pandoc emits these when Word source has a Heading style applied to an empty
# paragraph. Useless for RAG segmentation — drop them.
_EMPTY_HEAD_RE = re.compile(r"^#{1,6}\s*$")

# Plain-text 3GPP section number, e.g. '5.8.2.2 UE IP Address Management' or
# '5.3.3.4 Reception of the *RRCConnectionSetup* by the UE'.
# Anchored: full line, '<int>.' chain, title starts with [A-Z] (avoids prose
# like '5.3 GHz frequency …'), and allows markdown decoration chars (* _ [ \)
# inside the title so italicised parameter names don't break the match.
_PLAIN_SECT_RE = re.compile(
    r"^((?:\d+\.){1,5}\d+)\s+(\*?[A-Z][A-Za-z0-9 ,./&()\-*_\[\]\\]{1,150})$"
)

# 3GPP "Start of changes" / "End of changes" delimiters, in all the wild
# decoration variants seen in real CRs (asterisks, brackets, bold, dots).
_CHANGE_MARKER_RE = re.compile(
    r"(?i)(?:start|first|next|last|end)\s+of\s+changes?"
)


def _escape_fp_headings(lines: list[str]) -> int:
    """Escape '#3 (...)' false positives in-place. Returns count of fixes."""
    n = 0
    for i, line in enumerate(lines):
        if _FP_HEAD_RE.match(line):
            lines[i] = "\\" + line
            n += 1
    return n


def _drop_empty_headings(lines: list[str]) -> int:
    """Replace empty heading lines ('####  ') with a blank line in-place.
    Returns count dropped.
    """
    n = 0
    for i, line in enumerate(lines):
        if _EMPTY_HEAD_RE.match(line):
            lines[i] = ""
            n += 1
    return n


def _promote_plain_sections(lines: list[str]) -> int:
    """Promote plain-text 3GPP section numbers to '##' headings.

    Triggers when a line matches _PLAIN_SECT_RE AND is sandwiched between
    blank lines (real section titles in CRs are always standalone paragraphs;
    body references like 'see 5.3 Gbps throughput' are never line-start).

    Runs even when the file already has hash-headings — CRs frequently mix
    Heading-styled paragraphs (which pandoc emits as '#') with bold-only
    section titles (which arrive as plain text). The blank-blank sandwich
    keeps false-positives near zero.
    """
    n = 0
    last = len(lines) - 1
    for i, line in enumerate(lines):
        stripped = line.strip()
        if not _PLAIN_SECT_RE.match(stripped):
            continue
        prev_blank = (i == 0) or (lines[i-1].strip() == "")
        next_blank = (i == last) or (lines[i+1].strip() == "")
        if prev_blank and next_blank:
            lines[i] = "## " + stripped
            n += 1
    return n


def _shift_heading_depth(lines: list[str]) -> int:
    """Shift heading levels so the shallowest used becomes '#'. Returns shift."""
    min_level = None
    for L in lines:
        m = _REAL_HEAD_RE.match(L)
        if m:
            lv = len(m.group(1))
            if min_level is None or lv < min_level:
                min_level = lv
    if not min_level or min_level == 1:
        return 0
    shift = min_level - 1
    for i, line in enumerate(lines):
        m = _REAL_HEAD_RE.match(line)
        if m:
            lv = len(m.group(1))
            lines[i] = "#" * (lv - shift) + line[lv:]
    return shift


def _strip_cr_form_grid(lines: list[str]) -> int:
    """If a 3GPP 'Start of changes' marker is found at line N (with at least 50
    lines of content after it), drop the preceding ASCII grid table. The
    discarded region is summarised in a single-line note.

    Returns the number of lines dropped (0 if no marker / too little content).
    """
    start_idx = None
    for i, L in enumerate(lines):
        if _CHANGE_MARKER_RE.search(L):
            start_idx = i
            break
    if start_idx is None or start_idx < 5:
        return 0
    if len(lines) - start_idx < 50:
        return 0
    dropped = start_idx
    note = f"<!-- {dropped} lines of CR form metadata stripped (pre 'Start of changes') -->"
    lines[:start_idx] = [note, ""]
    return dropped


def normalize_headings(md: str) -> str:
    """Post-process pandoc-emitted markdown so RAG chunkers see clean section
    boundaries. Combines five fixes:
      1. Escape '#<digit>' false positives ('#3 (Illegal UE)' → '\\#3 (...)').
      2. Drop empty heading lines ('####  ' with no title) — Pandoc emits
         these for empty Word paragraphs that carry a Heading style. They
         create 0-length sections and confuse RAG chunkers.
      3. Strip the CR-form ASCII-grid metadata block when a 3GPP
         'Start of changes' delimiter exists below it.
      4. Promote plain-text section numbers ('5.8.2.2 ...') to '##' headings
         when the document has no ATX headings at all.
      5. Normalise heading depth so the shallowest level used becomes '#',
         removing per-doc drift where the same section type appears at '#',
         '##', or '###' across files.
    """
    lines = md.split("\n")
    _escape_fp_headings(lines)
    _drop_empty_headings(lines)
    _strip_cr_form_grid(lines)
    _promote_plain_sections(lines)
    _shift_heading_depth(lines)
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# EQUATION IMAGE OCR  (optional — requires pix2tex)
# ══════════════════════════════════════════════════════════════════════════════

# Images shorter than this (inches) are treated as inline equations, not figures.
# LibreOffice renders inline equations as tiny images (typically ~0.15–0.30 in tall).
_EQ_HEIGHT_MAX_IN = 0.5

# Regex to find inline image references produced by pandoc from DOCX media files.
# Matches: ![alt](media/imageN.ext){...attrs...}   (attrs may be absent)
_IMG_RE = re.compile(r'!\[([^\]]*)\]\(media/([^)]+)\)(\{[^}]*\})?')

# Regex to detect equation-sized images (height < 1 inch, expressed as 0.XXin)
_EQUATION_IMG_RE = re.compile(r'!\[[^\]]*\]\(media/[^)]+\)\{[^}]*height="0\.\d+in"')

_pix2tex_model = None  # module-level cache so the model is only loaded once


def _load_pix2tex():
    """Lazily load and cache the pix2tex LatexOCR model."""
    global _pix2tex_model
    if _pix2tex_model is None:
        try:
            from pix2tex.cli import LatexOCR
        except ImportError:
            raise ImportError(
                "pix2tex is required for equation OCR.\n"
                "Install it with:  pip install pix2tex"
            )
        print("Loading pix2tex model…", file=sys.stderr)
        _pix2tex_model = LatexOCR()
    return _pix2tex_model


def _extract_media(docx_path: str) -> dict:
    """Return {basename: bytes} for every file under word/media/ in a .docx."""
    media = {}
    with zipfile.ZipFile(docx_path, "r") as z:
        for name in z.namelist():
            if name.startswith("word/media/"):
                media[os.path.basename(name)] = z.read(name)
    return media


def _ocr_equation_image(img_bytes: bytes) -> str:
    """Run pix2tex on img_bytes; return LaTeX string or empty string on failure."""
    try:
        import PIL.Image
        model = _load_pix2tex()
        img = PIL.Image.open(io.BytesIO(img_bytes)).convert("RGB")
        return model(img)
    except Exception as exc:
        print(f"  ⚠ Equation OCR failed: {exc}", file=sys.stderr)
        return ""


def _replace_equation_images(md: str, media: dict) -> str:
    """
    Scan markdown for small inline image references (height < _EQ_HEIGHT_MAX_IN inches)
    and replace them with pix2tex LaTeX wrapped in $…$.

    Pandoc produces entries like:
        ![](media/image1.png){width="0.583in" height="0.167in"}
    for equations that LibreOffice converted to images.
    """
    if not media:
        return md

    n_replaced = 0
    n_failed   = 0

    def _repl(m):
        nonlocal n_replaced, n_failed
        full     = m.group(0)
        img_name = m.group(2)
        attrs    = m.group(3) or ""

        # ── size gate ─────────────────────────────────────────────────────────
        height_m = re.search(r'height="([\d.]+)in"', attrs)
        if height_m:
            if float(height_m.group(1)) >= _EQ_HEIGHT_MAX_IN:
                return full          # large → real figure, leave alone
        else:
            # No explicit size — fall back to actual pixel height via PIL
            img_bytes = media.get(img_name)
            if img_bytes:
                try:
                    import PIL.Image
                    h_px = PIL.Image.open(io.BytesIO(img_bytes)).height
                    # 96 DPI baseline: 0.5 in ≈ 48 px
                    if h_px >= 48:
                        return full
                except Exception:
                    pass

        img_bytes = media.get(img_name)
        if img_bytes is None:
            return full

        latex = _ocr_equation_image(img_bytes)
        if latex:
            n_replaced += 1
            return f"${latex}$"
        else:
            n_failed += 1
            return full

    result = _IMG_RE.sub(_repl, md)

    if n_replaced or n_failed:
        print(
            f"Equation OCR: {n_replaced} image(s) → LaTeX, {n_failed} failed",
            file=sys.stderr,
        )
    return result


# ══════════════════════════════════════════════════════════════════════════════
# EMBEDDED WORDML EQUATION EXTRACTION  (for .doc files)
# ══════════════════════════════════════════════════════════════════════════════
#
# Many .doc files saved by newer Word versions embed Word-2003 WordML XML
# snippets in the binary `Data` stream — one per inline math object. Each snippet
# contains a real `<m:oMath>` element (OOXML math). LibreOffice ignores these
# and renders equations as blank PNGs, but we can extract the WordML directly,
# wrap each in a minimal .docx, and let pandoc convert it to LaTeX.
#
# This only kicks in for .doc files where WordML math is present. If the Data
# stream has no <m:oMath>, extraction returns an empty list and the legacy
# blank-image / pix2tex path continues as before.

_DOC_XML_START = re.compile(rb'<\?xml version="1\.0"')
_DOC_XML_END   = re.compile(rb'</w:wordDocument>')
_OMATH_RE      = re.compile(r'<m:oMath[ >].*?</m:oMath>', re.DOTALL)

_AML_NS = "http://schemas.microsoft.com/aml/2001/core"

_EQ_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '</Types>'
)
_EQ_RELS_ROOT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/></Relationships>'
)
_EQ_DOC_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:aml="http://schemas.microsoft.com/aml/2001/core" '
    'xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint">'
    '<w:body><w:p>{MATH}</w:p></w:body></w:document>'
)


def _unwrap_aml_annotations(omath_xml: str) -> str:
    """
    Tracked-change annotations inside WordML math are wrapped as
    <aml:annotation><aml:content>…actual math…</aml:content></aml:annotation>.
    Pandoc ignores these wrappers and drops the inner text, producing empty
    equations. Unwrap them in-place before conversion.
    """
    wrapped = (
        '<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:aml="http://schemas.microsoft.com/aml/2001/core" '
        'xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint">'
        + omath_xml + '</root>'
    )
    root = etree.fromstring(wrapped.encode("utf-8"))
    for ann in list(root.iter(f"{{{_AML_NS}}}annotation")):
        parent = ann.getparent()
        if parent is None:
            continue
        idx = list(parent).index(ann)
        inner = []
        for content in ann.findall(f"{{{_AML_NS}}}content"):
            inner.extend(list(content))
        tail = ann.tail
        parent.remove(ann)
        for i, c in enumerate(inner):
            parent.insert(idx + i, c)
        if inner and tail:
            inner[-1].tail = (inner[-1].tail or "") + tail
    return etree.tostring(root[0], encoding="unicode")


def _omath_to_latex(omath_xml: str) -> str:
    """Wrap a single <m:oMath> in a minimal .docx and convert to LaTeX via pandoc."""
    try:
        omath_xml = _unwrap_aml_annotations(omath_xml)
    except Exception:
        pass  # fall through with original; pandoc may still produce partial text

    doc = _EQ_DOC_TEMPLATE.replace("{MATH}", omath_xml)
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
        with zipfile.ZipFile(tf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _EQ_CONTENT_TYPES)
            z.writestr("_rels/.rels", _EQ_RELS_ROOT)
            z.writestr("word/document.xml", doc)
        tmp = tf.name
    try:
        md = pypandoc.convert_file(tmp, "markdown", format="docx", extra_args=["--wrap=none"])
    finally:
        os.unlink(tmp)
    return md.strip()


def extract_doc_wordml_equations(doc_path: str) -> list:
    """
    Read a .doc file's OLE `Data` stream and return a list of LaTeX strings —
    one per inline equation, in document order.

    Returns an empty list if `olefile` is unavailable, the file has no Data
    stream, or no WordML math is embedded.
    """
    try:
        import olefile
    except ImportError:
        return []

    try:
        ole = olefile.OleFileIO(doc_path)
    except Exception:
        return []

    try:
        if not ole.exists("Data"):
            return []
        data = ole.openstream("Data").read()
    finally:
        ole.close()

    starts = [m.start() for m in _DOC_XML_START.finditer(data)]
    ends   = [m.end()   for m in _DOC_XML_END.finditer(data)]
    if not starts or not ends:
        return []

    blobs, j = [], 0
    for s in starts:
        while j < len(ends) and ends[j] <= s:
            j += 1
        if j < len(ends):
            blobs.append(data[s:ends[j]])

    # Dedup consecutively-identical blobs (Word writes each inline object twice),
    # preserving document order.
    seen, ordered = set(), []
    for b in blobs:
        h = hash(b)
        if h in seen:
            continue
        seen.add(h)
        ordered.append(b)

    latexes = []
    for blob in ordered:
        try:
            xml = blob.decode("utf-8", errors="replace")
        except Exception:
            continue
        m = _OMATH_RE.search(xml)
        if not m:
            continue
        try:
            latex = _omath_to_latex(m.group(0))
        except Exception as exc:
            print(f"  ⚠ WordML equation conversion failed: {exc}", file=sys.stderr)
            latex = ""
        latexes.append(latex)

    return latexes


def _replace_images_with_latex(md: str, latexes: list) -> str:
    """
    Replace each `![...](media/imageN.png){...}` in `md` with the next LaTeX
    string from `latexes`, in document order. Images without a corresponding
    entry (e.g. real figures outnumbering equations) are left alone.
    """
    it = iter(latexes)

    def _sub(m):
        try:
            latex = next(it)
        except StopIteration:
            return m.group(0)
        if not latex:
            return m.group(0)
        # pandoc emits the equation already wrapped with $…$
        return latex

    return _IMG_RE.sub(_sub, md)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 0: .DOC → .DOCX VIA LIBREOFFICE
# ══════════════════════════════════════════════════════════════════════════════

_SOFFICE_PATHS = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
    "/usr/bin/soffice",                                       # Linux
]


def _find_soffice():
    """Return the path to the LibreOffice soffice binary, or None."""
    for path in _SOFFICE_PATHS:
        if os.path.isfile(path):
            return path
    # Fall back to whatever is on PATH
    return shutil.which("soffice")


def doc_to_docx(source: str, outdir: str, timeout: int = 300) -> str:
    """
    Convert a .doc file to .docx using LibreOffice headless mode.

    Parameters
    ----------
    source  : path to the .doc file
    outdir  : directory where the .docx will be written
    timeout : seconds before the conversion is killed (default 300)

    Returns
    -------
    Path to the resulting .docx file.

    Raises
    ------
    FileNotFoundError : if LibreOffice is not installed
    RuntimeError      : if the conversion fails
    """
    soffice = _find_soffice()
    if soffice is None:
        raise FileNotFoundError(
            "LibreOffice not found. Install it with: brew install --cask libreoffice"
        )

    # Use a temporary user profile so concurrent runs don't clash
    with tempfile.TemporaryDirectory() as profile_dir:
        user_profile = f"-env:UserInstallation=file://{profile_dir}"
        cmd = [
            soffice,
            "--headless", "--norestore", "--invisible",
            "--nocrashreport", "--nodefault", "--nologo",
            "--nofirststartwizard",
            user_profile,
            "--convert-to", "docx",
            "--outdir", outdir,
            source,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)

        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (exit {result.returncode}): {result.stderr}"
            )

    # LibreOffice writes <basename>.docx in outdir
    base = os.path.splitext(os.path.basename(source))[0]
    docx_path = os.path.join(outdir, base + ".docx")

    if not os.path.isfile(docx_path):
        raise RuntimeError(
            f"LibreOffice ran but output file not found: {docx_path}\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}"
        )

    return docx_path


# ══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ══════════════════════════════════════════════════════════════════════════════

def convert(
    input_docx: str,
    output_md: str = None,
    md_format: str = "markdown",
    ocr_equations: bool = False,
    postprocess_headings: bool = True,
) -> str:
    """
    Full pipeline:
      0. If input is .doc, convert to .docx via LibreOffice first
      1. Inject markers into .docx XML (handles ins/del/move/fmt changes)
      2. Convert modified .docx → Markdown via pandoc
      3. (Optional) Replace equation images with pix2tex LaTeX
      4. Replace markers with <ins>/<del>/etc. HTML tags
      5. Clean up
      6. (Optional) Normalise headings for RAG-friendly section segmentation

    Parameters
    ----------
    input_docx           : path to the .docx or .doc file
    output_md            : path to write the .md file (optional)
    md_format            : pandoc output format — 'markdown', 'gfm', 'markdown_strict', etc.
    ocr_equations        : if True, run pix2tex on small equation images produced by
                           LibreOffice's .doc → .docx conversion (requires: pip install pix2tex)
    postprocess_headings : if True (default), apply RAG-friendly heading fixes
                           (escape '#<digit>' false positives, strip CR-form grid
                           before 'Start of changes', promote plain section
                           numbers, normalise heading depth)

    Returns
    -------
    Markdown string with track changes as HTML tags.
    """
    media: dict = {}
    doc_equations: list = []

    with tempfile.TemporaryDirectory() as tmpdir:
        # Step 0: convert .doc → .docx if needed
        if input_docx.lower().endswith(".doc"):
            # Extract WordML equations from the .doc's OLE Data stream BEFORE
            # LibreOffice drops them. These give us real LaTeX for inline math.
            print("Extracting embedded WordML equations from .doc…", file=sys.stderr)
            doc_equations = extract_doc_wordml_equations(input_docx)
            if doc_equations:
                print(f"  Found {len(doc_equations)} inline equation(s).", file=sys.stderr)

            print("Converting .doc → .docx via LibreOffice…", file=sys.stderr)
            input_docx = doc_to_docx(input_docx, tmpdir)
            print(f"Converted: {input_docx}", file=sys.stderr)

        # Pre-extract media into memory before tmpdir is torn down.
        # (The tmpdir is deleted when the `with` block exits; bytes stay in memory.)
        if ocr_equations:
            media = _extract_media(input_docx)
            if media:
                print(
                    f"Found {len(media)} media file(s); will OCR small equation images with pix2tex.",
                    file=sys.stderr,
                )

        marked = os.path.join(tmpdir, "marked.docx")
        create_marked_docx(input_docx, marked)

        md = pypandoc.convert_file(
            marked,
            md_format,
            format="docx",
            extra_args=["--wrap=none", "--markdown-headings=atx"],
        )

    # Step 3a: splice WordML-extracted LaTeX into image placeholders (for .doc)
    if doc_equations:
        md = _replace_images_with_latex(md, doc_equations)

    # Step 3b (optional): any remaining images get the pix2tex path
    if media:
        md = _replace_equation_images(md, media)
    elif not ocr_equations and not doc_equations and _EQUATION_IMG_RE.search(md):
        print(
            "⚠  Output contains formula images (equations rendered as PNG by LibreOffice).\n"
            "   Re-run with --ocr-equations to convert them to LaTeX text via pix2tex.",
            file=sys.stderr,
        )

    md = markers_to_tags(md)
    md = clean_markdown(md)

    if postprocess_headings:
        md = normalize_headings(md)

    if output_md:
        with open(output_md, "w", encoding="utf-8") as f:
            f.write(md)

    return md


# ══════════════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert .docx (or .doc) with track changes to Markdown with <ins>/<del> tags."
    )
    parser.add_argument("input", help="Path to the .docx or .doc file")
    parser.add_argument("output", nargs="?", default=None, help="Path to write the .md file (default: input with .md extension)")
    parser.add_argument(
        "-f", "--format",
        default="markdown",
        help="Pandoc output format: markdown (default), gfm, markdown_strict, etc.",
    )
    parser.add_argument(
        "--ocr-equations",
        action="store_true",
        default=False,
        help=(
            "OCR equation images produced by LibreOffice's .doc → .docx conversion "
            "and replace them with LaTeX text. Requires: pip install pix2tex"
        ),
    )
    parser.add_argument(
        "--no-heading-fix",
        action="store_true",
        default=False,
        help=(
            "Disable RAG-friendly heading post-processing (escape '#<digit>' "
            "false positives, strip CR-form grid, promote plain section numbers, "
            "normalise heading depth). Default: on."
        ),
    )

    args = parser.parse_args()
    src = args.input
    dst = args.output if args.output else src.rsplit(".", 1)[0] + ".md"

    result = convert(
        src,
        dst,
        md_format=args.format,
        ocr_equations=args.ocr_equations,
        postprocess_headings=not args.no_heading_fix,
    )
    print(result)