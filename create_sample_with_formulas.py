"""
create_sample_with_formulas.py

Creates a minimal .docx with:
  - Regular paragraph text
  - Track-change insertions and deletions
  - Inline OOXML math equations (<m:oMath>)

Then converts that .docx to .doc via LibreOffice so we can compare
how the two formats are handled by docx_parsing.py.

Usage:
    python create_sample_with_formulas.py
Outputs:
    input/sample_formulas.docx   ← equations as OOXML math (pandoc → LaTeX)
    input/sample_formulas.doc    ← equations as OLE images after LO conversion
"""

import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

# ── OOXML namespaces ──────────────────────────────────────────────────────────
W_NS   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS   = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WPKG   = "http://schemas.openxmlformats.org/package/2006/relationships"

# ── Minimal .docx skeleton ───────────────────────────────────────────────────

CONTENT_TYPES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
"""

ROOT_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>
"""

WORD_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
</Relationships>
"""

STYLES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>
"""

SETTINGS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:trackChanges/>
</w:settings>
"""

# ── document.xml ─────────────────────────────────────────────────────────────
# Contains:
#   Para 1 — plain text intro
#   Para 2 — track-change insertion ("proposed amendment")
#   Para 3 — track-change deletion ("obsolete clause")
#   Para 4 — inline equation: E = mc²  (simple run-based math)
#   Para 5 — inline equation: fraction x/2  using <m:f>
#   Para 6 — equation inside a track-change insertion
#   Para 7 — equation inside a track-change deletion

DOCUMENT_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>

    <!-- Para 1: plain intro -->
    <w:p>
      <w:r><w:t>This document tests formula handling in track-change pipelines.</w:t></w:r>
    </w:p>

    <!-- Para 2: track-change INSERTION of plain text -->
    <w:p>
      <w:r><w:t xml:space="preserve">The bandwidth is limited to </w:t></w:r>
      <w:ins w:id="1" w:author="Alice" w:date="2025-01-01T00:00:00Z">
        <w:r><w:t>100 MHz</w:t></w:r>
      </w:ins>
      <w:r><w:t xml:space="preserve"> per channel.</w:t></w:r>
    </w:p>

    <!-- Para 3: track-change DELETION of plain text -->
    <w:p>
      <w:r><w:t xml:space="preserve">The old limit was </w:t></w:r>
      <w:del w:id="2" w:author="Bob" w:date="2025-01-02T00:00:00Z">
        <w:r><w:delText>50 MHz</w:delText></w:r>
      </w:del>
      <w:r><w:t xml:space="preserve"> per channel.</w:t></w:r>
    </w:p>

    <!-- Para 4: inline equation E = mc^2 using simple <m:oMath> run -->
    <w:p>
      <w:r><w:t xml:space="preserve">Einstein&#8217;s equation </w:t></w:r>
      <m:oMath>
        <m:r><m:t>E</m:t></m:r>
        <m:r><m:t>=</m:t></m:r>
        <m:r><m:t>m</m:t></m:r>
        <m:r><m:t>c</m:t></m:r>
        <m:sSup>
          <m:e><m:r><m:t>c</m:t></m:r></m:e>
          <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSup>
      </m:oMath>
      <w:r><w:t xml:space="preserve"> describes mass-energy equivalence.</w:t></w:r>
    </w:p>

    <!-- Para 5: inline fraction x / 2 using <m:f> -->
    <w:p>
      <w:r><w:t xml:space="preserve">The average is </w:t></w:r>
      <m:oMath>
        <m:f>
          <m:fPr><m:type m:val="bar"/></m:fPr>
          <m:num><m:r><m:t>x</m:t></m:r></m:num>
          <m:den><m:r><m:t>2</m:t></m:r></m:den>
        </m:f>
      </m:oMath>
      <w:r><w:t xml:space="preserve"> across all samples.</w:t></w:r>
    </w:p>

    <!-- Para 6: equation INSERTED via track change -->
    <w:p>
      <w:r><w:t xml:space="preserve">The proposed formula is </w:t></w:r>
      <w:ins w:id="3" w:author="Alice" w:date="2025-01-03T00:00:00Z">
        <m:oMath>
          <m:r><m:t>&#x3B1;</m:t></m:r>
          <m:r><m:t>+</m:t></m:r>
          <m:r><m:t>&#x3B2;</m:t></m:r>
          <m:r><m:t>=</m:t></m:r>
          <m:r><m:t>&#x3B3;</m:t></m:r>
        </m:oMath>
      </w:ins>
      <w:r><w:t xml:space="preserve"> where alpha and beta are inputs.</w:t></w:r>
    </w:p>

    <!-- Para 7: equation DELETED via track change -->
    <w:p>
      <w:r><w:t xml:space="preserve">The old formula </w:t></w:r>
      <w:del w:id="4" w:author="Bob" w:date="2025-01-04T00:00:00Z">
        <m:oMath>
          <m:r><m:t>&#x3B1;</m:t></m:r>
          <m:r><m:t>=</m:t></m:r>
          <m:r><m:t>0</m:t></m:r>
        </m:oMath>
      </w:del>
      <w:r><w:t xml:space="preserve"> is no longer valid.</w:t></w:r>
    </w:p>

    <w:sectPr/>
  </w:body>
</w:document>
"""


def build_docx(out_path: str):
    """Write a minimal .docx ZIP at out_path."""
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",       CONTENT_TYPES)
        zf.writestr("_rels/.rels",               ROOT_RELS)
        zf.writestr("word/_rels/document.xml.rels", WORD_RELS)
        zf.writestr("word/document.xml",         DOCUMENT_XML)
        zf.writestr("word/styles.xml",           STYLES)
        zf.writestr("word/settings.xml",         SETTINGS)
    print(f"Created: {out_path}")


def docx_to_doc(docx_path: str, out_dir: str) -> str:
    """Convert .docx → .doc via LibreOffice headless."""
    soffice = shutil.which("soffice")
    if soffice is None:
        candidates = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/usr/bin/soffice",
        ]
        for c in candidates:
            if os.path.isfile(c):
                soffice = c
                break
    if soffice is None:
        print("LibreOffice not found — skipping .doc creation.", file=sys.stderr)
        return ""

    with tempfile.TemporaryDirectory() as profile_dir:
        user_profile = f"-env:UserInstallation=file://{profile_dir}"
        cmd = [
            soffice,
            "--headless", "--norestore", "--invisible",
            "--nocrashreport", "--nodefault", "--nologo",
            "--nofirststartwizard",
            user_profile,
            "--convert-to", "doc",
            "--outdir", out_dir,
            docx_path,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

    base = os.path.splitext(os.path.basename(docx_path))[0]
    doc_path = os.path.join(out_dir, base + ".doc")
    if os.path.isfile(doc_path):
        print(f"Created: {doc_path}")
        return doc_path
    else:
        print(f"LibreOffice error: {result.stderr}", file=sys.stderr)
        return ""


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.join(here, "input")
    os.makedirs(input_dir, exist_ok=True)

    docx_out = os.path.join(input_dir, "sample_formulas.docx")
    build_docx(docx_out)

    doc_out = docx_to_doc(docx_out, input_dir)
    if doc_out:
        print("\nBoth files ready:")
        print(f"  {docx_out}")
        print(f"  {doc_out}")
    else:
        print("\nOnly .docx created (LibreOffice unavailable for .doc).")
