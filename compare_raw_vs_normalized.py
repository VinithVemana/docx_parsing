"""Side-by-side comparison of raw pandoc output vs RAG-normalized output.

Usage:
    python compare_raw_vs_normalized.py                       # default dirs
    python compare_raw_vs_normalized.py --only-promoted       # files where promotion fired
    python compare_raw_vs_normalized.py --limit 20 --out audit.tsv

For each markdown file present in both raw and normalized output dirs, counts:
  - hash-headings (real ATX)
  - empty heading lines ('#### ')
  - false-positive '#<digit>' lines
  - first hash-heading line number

Writes a TSV report so reviewers can sort/filter the diff.
"""
from __future__ import annotations
import argparse
import re
import sys
from pathlib import Path
from collections import Counter

REAL_HEAD_RE  = re.compile(r"^#{1,6}\s+\S")
EMPTY_HEAD_RE = re.compile(r"^#{1,6}\s*$")
FP_RE         = re.compile(r"^#\d")
GRID_NOTE_RE  = re.compile(r"^<!-- \d+ lines of CR form metadata stripped")


def stats(text: str) -> dict:
    lines = text.split("\n")
    real_idx = [i for i, L in enumerate(lines) if REAL_HEAD_RE.match(L)]
    empty_idx = [i for i, L in enumerate(lines) if EMPTY_HEAD_RE.match(L)]
    fp_idx = [i for i, L in enumerate(lines) if FP_RE.match(L)]
    return {
        "headings":  len(real_idx),
        "empty":     len(empty_idx),
        "fp":        len(fp_idx),
        "first_at":  real_idx[0] + 1 if real_idx else -1,
        "grid_note": bool(GRID_NOTE_RE.match(lines[0])) if lines else False,
        "size":      len(text),
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--raw",  default="output/change_requests_raw_pandoc",
                    help="Dir with raw pandoc output (--no-heading-fix)")
    ap.add_argument("--norm", default="output/change_requests",
                    help="Dir with RAG-normalised output")
    ap.add_argument("--only-promoted", action="store_true",
                    help="Only emit files where promotion fired (norm has more headings)")
    ap.add_argument("--limit", type=int, default=0,
                    help="Stop after N matching files (default: all)")
    ap.add_argument("--out", default=None,
                    help="Write TSV to this path (default: stdout)")
    args = ap.parse_args()

    raw_dir  = Path(args.raw)
    norm_dir = Path(args.norm)
    if not raw_dir.is_dir() or not norm_dir.is_dir():
        sys.exit(f"Missing dir: {raw_dir} or {norm_dir}")

    headers = [
        "file",
        "raw_heads", "norm_heads", "delta",
        "raw_empty", "norm_empty",
        "raw_fp", "norm_fp",
        "raw_first_at", "norm_first_at",
        "grid_stripped",
    ]
    out_lines = ["\t".join(headers)]

    diff_buckets = Counter()
    common = sorted(p.name for p in raw_dir.glob("*.md") if (norm_dir / p.name).is_file())

    n_emitted = 0
    for name in common:
        raw_txt  = (raw_dir / name).read_text(errors="replace")
        norm_txt = (norm_dir / name).read_text(errors="replace")
        s_raw  = stats(raw_txt)
        s_norm = stats(norm_txt)
        delta_heads = s_norm["headings"] - s_raw["headings"]

        if args.only_promoted and delta_heads <= 0:
            continue

        diff_buckets["files"] += 1
        diff_buckets["delta_heads"] += delta_heads
        diff_buckets["delta_empty"] += s_raw["empty"] - s_norm["empty"]
        diff_buckets["delta_fp"]    += s_raw["fp"]    - s_norm["fp"]
        diff_buckets["grid_stripped"] += int(s_norm["grid_note"])

        row = [
            name,
            s_raw["headings"], s_norm["headings"], delta_heads,
            s_raw["empty"], s_norm["empty"],
            s_raw["fp"], s_norm["fp"],
            s_raw["first_at"], s_norm["first_at"],
            "Y" if s_norm["grid_note"] else "",
        ]
        out_lines.append("\t".join(str(x) for x in row))
        n_emitted += 1
        if args.limit and n_emitted >= args.limit:
            break

    body = "\n".join(out_lines) + "\n"
    if args.out:
        Path(args.out).write_text(body)
        print(f"wrote {n_emitted} rows → {args.out}", file=sys.stderr)
    else:
        sys.stdout.write(body)

    # aggregate summary to stderr so reviewer sees overall delta
    print("── aggregate ──", file=sys.stderr)
    for k, v in diff_buckets.items():
        print(f"  {k:>14}: {v}", file=sys.stderr)


if __name__ == "__main__":
    main()
