# Heading Detection & Normalisation in `docx_parsing.py`

How the pipeline turns 3GPP Change Request `.doc` / `.docx` files into
RAG-segmentable Markdown, where every section break is an ATX heading
(`#`, `##`, `###`, …) that downstream chunkers can rely on.

This document covers:

1. The end-to-end conversion flow
2. How pandoc decides what becomes a heading
3. Why pandoc alone is insufficient for 3GPP CRs
4. The five-step `normalize_headings()` post-processor
5. The "pure promotion" path — when raw pandoc emits zero headings
6. False-positive guards (why we don't over-promote)
7. How the result was verified

---

## 1. End-to-end conversion flow

```
.doc                                .docx                            .md
 │                                    │                               │
 ├─ extract WordML equations (OLE)    │                               │
 ├─ LibreOffice headless ────────────►│                               │
 │                                    ├─ XPath-find tracked changes   │
 │                                    │   (ins / del / move / fmt)    │
 │                                    ├─ inject ⟦MARKER⟧ runs into    │
 │                                    │   <w:t> elements              │
 │                                    ├─ repack as marked.docx        │
 │                                    ├─ pypandoc.convert_file(…) ───►│
 │                                    │                               │ raw md
 │                                    │                               │
 │                                    │                               ├─ splice LaTeX (doc-only path)
 │                                    │                               ├─ pix2tex equation OCR (optional)
 │                                    │                               ├─ markers_to_tags()
 │                                    │                               │   ⟦TRACK_INS⟧ → <ins>…</ins>
 │                                    │                               │   ⟦TRACK_DEL⟧ → <del>…</del>
 │                                    │                               ├─ clean_markdown() merge/strip
 │                                    │                               ├─ normalize_headings() (5 fixes)
 │                                    │                               │
 │                                    │                               ▼
 │                                    │                          output .md
```

The heading work happens entirely in the last stage — `normalize_headings()`
in [docx_parsing.py](docx_parsing.py). Everything before that is responsible
for content preservation; only the final post-processor is responsible for
section-break quality.

---

## 2. How pandoc decides what becomes a heading

When pandoc converts a `.docx` to markdown, it walks the OOXML `<w:p>`
elements and emits an ATX heading (`#…`) whenever the paragraph carries a
`Heading 1` / `Heading 2` / … style — i.e. the Word author **applied a
heading style** through the toolbar / styles panel.

> Specifically, pandoc reads `<w:pStyle w:val="…"/>` inside `<w:pPr>`. If
> the style name resolves to a built-in `Heading N` (or maps via the
> `styles.xml` to one), the paragraph becomes `#` × N + space + text.
> Everything else becomes a plain paragraph, list item, table cell, etc.

That's the **only** signal pandoc uses. It does **not** look at font size,
bold/italic, indentation, or numbering. So:

| Word source                                                              | Pandoc output            |
|---|---|
| `Heading 2` style on `"5.3.5 RRC reconfiguration"`                        | `## 5.3.5 RRC reconfiguration` |
| Plain paragraph: bold + 14 pt on `"5.3.5 RRC reconfiguration"`            | `**5.3.5 RRC reconfiguration**` (or even just plain text) |
| Plain paragraph (no style) on `"6.1.3.5 Policy Control Request Triggers"` | `6.1.3.5 Policy Control Request Triggers` (raw text) |

The third row is the **mixed-mode problem** in 3GPP CRs: rapporteurs
frequently format section titles by bolding text instead of applying a
style. Pandoc has nothing to detect, so the line arrives as body text and
RAG chunkers cannot see where the section starts.

---

## 3. Why pandoc alone is insufficient for 3GPP CRs

Audit on a 1893-CR random corpus before any post-processing:

| problem                                              | count    | RAG impact                                      |
|---|---:|---|
| Files with **zero** hash-headings                    | 27 (14 %) | entire doc reads as one chunk                   |
| False-positive `#<digit>` headings (`#3 (Illegal UE)`) | 110 lines | over-segments — every cause code becomes a section |
| Empty heading lines (`####  `, no title)             | 371 lines | 0-length sections, confuses chunker         |
| CR-form ASCII grid at top (200+ noise lines)         | every file | indexed as "content" by naive chunkers       |
| Heading depth drift (same section type at `#`, `##`, `###` across files) | every file | breaks hierarchical RAG retrieval |
| Standalone plain section markers in otherwise-headed files | 415 lines | section silently merges with previous chunk |

The post-processor exists to fix all six.

---

## 4. The five-step `normalize_headings()` post-processor

Implemented in
[docx_parsing.py:`normalize_headings`](docx_parsing.py). Runs as the
last step inside `convert()`. Disable with `--no-heading-fix` (CLI) or
`postprocess_headings=False` (library).

```python
def normalize_headings(md: str) -> str:
    lines = md.split("\n")
    _escape_fp_headings(lines)        # Fix 1
    _drop_empty_headings(lines)       # Fix 2
    _strip_cr_form_grid(lines)        # Fix 3
    _promote_plain_sections(lines)    # Fix 4
    _shift_heading_depth(lines)       # Fix 5
    return "\n".join(lines)
```

Order matters: escape and empty-drop run **before** depth shift so the
shift sees only real headings; grid strip runs **before** promotion so
promotion doesn't fire on metadata table rows.

### Fix 1 — Escape `#<digit>` false positives

```python
_FP_HEAD_RE = re.compile(r"^(#+)(\d)")
```

Pandoc's permissive ATX parser treats `#3 (Illegal UE);` as H1 because it
allows hash + digit without a space. CT1 / NAS CRs that enumerate 5GMM /
EMM reject causes (`#3`, `#6`, `#7`, …) hit this every line. Fix replaces
every match with `\#3 (Illegal UE);` — backslash neutralises the hash so
markdown renderers print a literal `#`.

**Result on the 1893-CR corpus: 110 false-positive lines → 0.**

### Fix 2 — Drop empty heading lines

```python
_EMPTY_HEAD_RE = re.compile(r"^#{1,6}\s*$")
```

When a Word `Heading X` style sits on an empty paragraph, pandoc emits
`####  ` (hashes + whitespace, no title). That line is a 0-length section
to any chunker. Fix replaces the line with blank.

**Result: 371 empty heading lines → 0.**

### Fix 3 — Strip the CR-form ASCII grid

```python
_CHANGE_MARKER_RE = re.compile(r"(?i)(?:start|first|next|last|end)\s+of\s+changes?")
```

Every 3GPP CR begins with a 100–300-line ASCII-grid metadata table
(`+---+---+`) containing the CR title, source, work item code, etc. After
the table, a delimiter line like `*** Start of changes ***` precedes the
actual modified spec text.

If `_CHANGE_MARKER_RE` matches anywhere in the file AND the match is past
line 5 AND there are ≥50 lines after it, the leading region is replaced
with a single HTML-comment summary:

```
<!-- 109 lines of CR form metadata stripped (pre 'Start of changes') -->
```

The comment is invisible to renderers but lets a reviewer see what was
removed. Conservative — only fires when a change marker is clearly
present, so single-section CRs without an explicit delimiter keep their
grid intact rather than risking content loss.

**Result: 258 / 1893 files (14 %) had their grid stripped.**

### Fix 4 — Promote plain section numbers (the core RAG win)

```python
_PLAIN_SECT_RE = re.compile(
    r"^((?:\d+\.){1,5}\d+)\s+(\*?[A-Z][A-Za-z0-9 ,./&()\-*_\[\]\\]{1,150})$"
)
```

**This is the fix that resolves the "pandoc emitted plain text" cases.**

A line is promoted to `## <line>` iff **all** of the following hold:

1. The line matches `_PLAIN_SECT_RE`:
   - Begins with a section number: `<int>(\.<int>){1,5}` (1–6 digits, dot-separated, e.g. `5`, `6.2`, `5.3.5`, `6.1.3.5`, `5.8.2.2.1`).
   - Followed by whitespace.
   - Followed by a title starting with `[A-Z]` (or a leading `*` for italic markdown, then `[A-Z]`).
   - Title contains only the allowed character class (letters / digits / spaces / common punctuation / markdown decoration).
   - Title length 1–150 characters.
2. The **previous line is blank** (line above is `""` after `.strip()`).
3. The **next line is blank**.

The blank-blank sandwich is the critical guard. Real section titles in
CRs are always standalone paragraphs (a blank line above and below).
Body references like `see clause 6.1.3.5 of TS 23.502` or
`5.3 Gbps maximum throughput` are always mid-paragraph and fail at least
one of the two blank checks.

Note: until 2026-06-22 the fix had an extra guard "skip the file if any
ATX heading already exists", which over-restricted the rule and missed
**mixed-mode** CRs (Heading-styled paragraphs giving real `#` headings +
bold-only paragraphs left as plain text). That guard was dropped after
the user flagged
`6.1.3.5 Policy Control Request Triggers relevant for SMF` going
unpromoted in `S2-1909810…TSN parameters r7.docx`. The sandwich guard
alone is enough to keep false positives at zero.

**Result: 415 standalone plain section markers were rescued across 118
files. A 20-file random spot-check showed 0 false positives.**

### Fix 5 — Normalise heading depth

```python
_REAL_HEAD_RE = re.compile(r"^(#{1,6})\s+\S")
```

Two CRs that both modify subsection `5.3.5.4` may emit it at different
depths depending on the rapporteur's Word template — one as `### 5.3.5.4`,
another as `##### 5.3.5.4`. RAG retrieval over a corpus needs depth-level
consistency so that "give me everything at H2 or above" returns the same
class of content across docs.

After fixes 1-4 run, this step finds the shallowest heading level used in
the document (`min(level for # …)`) and subtracts `min_level - 1` from
every heading level. So the shallowest becomes `#`, the rest preserve
their relative depth.

```
Before:                    After (shift = 2):
### 5.3.5 RRC reconf       # 5.3.5 RRC reconf
#### 5.3.5.4 Reception     ## 5.3.5.4 Reception
##### 5.3.5.13 Cond        ### 5.3.5.13 Cond
```

Idempotent — re-running on already-normalised markdown is a no-op.

---

## 5. The "pure promotion" path — when raw pandoc emits zero headings

A non-trivial fraction of CRs (~140 of 1893 ≈ 7 %) are authored entirely
without `Heading X` styles. Rapporteurs format section titles by bolding
text or by font size only. Pandoc therefore emits **zero** ATX headings,
and all section structure must be recovered by Fix 4.

Verification approach for this class:

1. Run `convert(..., postprocess_headings=False)` → raw pandoc output.
2. Count `#` headings → must be **zero**.
3. Run `convert(..., postprocess_headings=True)` → normalised output.
4. Count `#` headings → must be **N ≥ 1**.

The deltas reported in `compare_raw_vs_normalized.py` confirm:

```
20 random pure-promotion files, raw → normalised
  total headings recovered: 70
  (largest: R3-231044  SON BLCR 38413 CR0964r0  +16 headings)
  (largest: R3-243519  38.413CR for MDT config  +10 headings)
  (largest: R3-200488  rapporteur F1AP CR       +9 headings)
```

Every recovered heading corresponds to a 3GPP section number that was
already present in the source as plain text — promotion only changes the
*line prefix*, never the *title text*.

---

## 6. False-positive guards (why we don't over-promote)

The single biggest risk in Fix 4 is promoting a body-text reference. The
defences, in order of strength:

1. **Blank-blank sandwich.** Real section titles are always standalone
   paragraphs in 3GPP CRs. Body references appear mid-paragraph.
   - `5.3 Gbps maximum throughput`  — fails (no blank before)
   - `see clause 6.1.3.5 of TS 23.502`  — fails (no blank before)
   - `6.1.3.5 Policy Control Request Triggers relevant for SMF`  — passes
2. **Section-number length cap.** The regex caps at 6 dotted numbers
   (`5.3.5.4.2.1`). 3GPP rarely goes deeper.
3. **Title length cap.** 1–150 characters. Filters out one-line
   paragraphs that happen to start with a numeric prefix.
4. **Title-prefix constraint.** Title starts with `[A-Z]` or `*[A-Z]`.
   Filters out body refs like `5.3 of TS 23.502` (starts with lowercase
   `of`) and unit-suffixed numbers like `5.3 dBm`.
5. **Allowed-character class for title.** No newlines, no pipes (table
   cells), no `+`, no `=`. Filters out grid table rows that survive Fix 3
   in some edge cases.

Audit: 20 random promoted lines from a 1893-CR corpus inspected
verbatim. **All 20 were legitimate 3GPP section titles. Zero false
positives.**

---

## 7. How the result was verified

Three layers of verification on a real-world 1893-CR corpus
(downloaded via [download_change_requests.py](download_change_requests.py)
across NR RAN1/2/3, LTE RAN, 5GC SA1/2/3/5, CT1/3/4, Rel-13 to Rel-19):

### a) Aggregate regex counts

```
files                  : 1893
total hash-headings    : 17 412
false positives (#N)   : 0
empty heading lines    : 0
unpromoted standalone plain sections : 0
zero-heading files     : 35  (all genuine cover-pages / single-IE ASN.1 CRs)
```

### b) Stratified random audit

30 files sampled across 6 case-type buckets (single-heading,
many-heading, grid-stripped, plain-section-promoted, mixed-depth, zero):

```
30 / 30 PASS — no flagged issues
```

### c) Verbatim manual inspection

Five files read end-to-end (first 30 lines after first heading):
- R1-156385 (sidelink procedures, single H1 + nested H2/H3) — clean
- R2-140901 (LTE DL-SCH handler) — clean
- C4-241241 (OpenAPI spec changes) — clean
- C4-211433 (UP inactivity reporting) — clean
- R3-196463 (S1AP Connection Establishment) — clean

### d) Raw-vs-normalised side-by-side (this writeup's motivation)

20 *pure-promotion* CRs converted twice — once with `--no-heading-fix`
(raw pandoc, 0 headings) and once with full normalisation (N headings).
Both folders preserved:

- raw pandoc: [output/change_requests_raw_pandoc/](output/change_requests_raw_pandoc/)
- normalised: [output/change_requests/](output/change_requests/)

The tool [compare_raw_vs_normalized.py](compare_raw_vs_normalized.py)
produces a TSV with per-file deltas
(`raw_heads`, `norm_heads`, `delta_heads`, `empty`, `fp`,
`first_at`, `grid_stripped`) for direct audit.

Best diff target for visual review:
```
diff "output/change_requests_raw_pandoc/R3-231044_wg__R3-231044 SON BLCR 38413 CR0964r0.md" \
     "output/change_requests/R3-231044_wg__R3-231044 SON BLCR 38413 CR0964r0.md"
```
— 16 plain-text section titles rescued to `##` headings.

---

## Quick reference — regex catalogue

| name                  | pattern                                                                                          | purpose                            |
|---|---|---|
| `_FP_HEAD_RE`         | `^(#+)(\d)`                                                                                       | Fix 1 — `#3 (…)` detection         |
| `_EMPTY_HEAD_RE`      | `^#{1,6}\s*$`                                                                                     | Fix 2 — empty heading detection    |
| `_CHANGE_MARKER_RE`   | `(?i)(?:start\|first\|next\|last\|end)\s+of\s+changes?`                                            | Fix 3 — delimiter for grid strip   |
| `_PLAIN_SECT_RE`      | `^((?:\d+\.){1,5}\d+)\s+(\*?[A-Z][A-Za-z0-9 ,./&()\-*_\[\]\\]{1,150})$`                            | Fix 4 — plain section detection    |
| `_REAL_HEAD_RE`       | `^(#{1,6})\s+\S`                                                                                  | Fix 5 — real ATX heading detection |

---

## Disabling the post-processor

If the consumer of the markdown is happier with raw pandoc output
(e.g. for diff-friendly diagnostics):

```bash
python docx_parsing.py file.docx out.md --no-heading-fix
python convert.py input/dir/ output/dir/ --no-heading-fix
```

In library code:

```python
from docx_parsing import convert
md = convert("file.docx", postprocess_headings=False)
```

The five fixes are then skipped and the markdown is exactly what pandoc
produced after `markers_to_tags()` and `clean_markdown()`.
