"""
Download random 3GPP Change Request files (WG TDoc + TSG TDoc) for testing the
docx_parsing pipeline on a wide variety of CRs.

Source: https://portal.3gpp.org/Home.aspx#/55932-change-requests

Pipeline:
  1. Open the Change Requests page with Playwright.
  2. For each --spec (or the default popular set), submit the on-page search,
     scrape the rgCrList grid (WG TDoc # + TSG TDoc # contribution UIDs).
  3. Randomly sample --n contribution UIDs across all collected rows.
  4. For each UID, GET DownloadTDoc.aspx, extract the redirect URL from the
     inline `window.location.href='...zip'` script, download the .zip via
     requests, and extract any .doc/.docx files into the output dir.

Usage:
    python download_change_requests.py                           # default: 15 random CRs across popular specs
    python download_change_requests.py --n 25                    # 25 random CRs
    python download_change_requests.py --specs 38.331 23.501     # restrict to listed specs
    python download_change_requests.py --kinds wg                # download only WG TDoc files (skip TSG)
    python download_change_requests.py --kinds tsg               # download only TSG TDoc files
    python download_change_requests.py --output input/crs/       # custom output dir (default: input/change_requests/)
    python download_change_requests.py --keep-zip                # also keep the original .zip alongside extracted files
    python download_change_requests.py --headed                  # show browser (debug)
"""

from __future__ import annotations

import argparse
import io
import random
import re
import sys
import time
import zipfile
from pathlib import Path
from urllib.parse import urlparse

import requests
from playwright.sync_api import sync_playwright
from tqdm import tqdm

PORTAL_URL = "https://portal.3gpp.org/Home.aspx#/55932-change-requests"
DOWNLOAD_URL = "http://portal.3gpp.org/ngppapp/DownloadTDoc.aspx?contributionUid={uid}"
SPEC_INPUT_ID = "dnn_ctr559_View_ctl00_ctl02_ctr600_ChangeRequestList_rpbCrSearch_i0_txtSpecificationNumber"
SEARCH_BTN_ID = "dnn_ctr559_View_ctl00_ctl02_ctr600_ChangeRequestList_rpbCrSearch_i0_btnSearch"

# Mixed series for variety (NR RAN1/2/3, NAS, SA2 5GC, LTE)
DEFAULT_SPECS = [
    "38.331",  # NR RRC
    "38.211",  # NR PHY general
    "38.213",  # NR PHY procedures (control)
    "38.321",  # NR MAC
    "36.331",  # LTE RRC
    "23.501",  # 5G system architecture (SA2)
    "24.501",  # 5GS NAS protocol (CT1)
]

REDIRECT_RE = re.compile(r"window\.location\.href\s*=\s*['\"]([^'\"]+)['\"]")


def collect_uids_for_spec(page, spec: str) -> list[dict]:
    """Run the on-page CR search for one spec; return [{wg_uid, tsg_uid, spec, cr, rel, title}, ...].

    Uses the visible grid header row to map column names to indices, so the
    parsing survives small layout changes.
    """
    page.evaluate(
        """(args) => {
            const inp = document.querySelector('#' + args.specId);
            inp.value = args.spec;
            document.querySelector('#' + args.btnId).click();
        }""",
        {"specId": SPEC_INPUT_ID, "btnId": SEARCH_BTN_ID, "spec": spec},
    )
    # Wait until the grid is reloaded — body row spec cell (index 2) matches the requested spec
    deadline = time.monotonic() + 90
    while time.monotonic() < deadline:
        state = page.evaluate(
            """(spec) => {
                const grid = document.querySelector('#rgCrList');
                if (!grid) return {ready: false};
                const rows = Array.from(grid.querySelectorAll('tbody tr')).filter(r => r.querySelectorAll('td').length >= 10);
                if (rows.length === 0) return {ready: false, rows: 0};
                const specCell = rows[0].querySelectorAll('td')[2]?.innerText.trim();
                return {ready: specCell === spec, rows: rows.length, specCell};
            }""",
            spec,
        )
        if state and state.get("ready"):
            break
        time.sleep(1)
    time.sleep(1)  # final settle

    rows = page.evaluate(
        """() => {
            const grid = document.querySelector('#rgCrList');
            if (!grid) return [];
            return Array.from(grid.querySelectorAll('tbody tr'))
                .filter(r => r.querySelectorAll('td').length >= 10)
                .map(r => ({
                    cells: Array.from(r.querySelectorAll('td')).map(c => c.innerText.trim()),
                    tdocLinks: Array.from(r.querySelectorAll('a[href*="DownloadTDoc"]')).map(a => a.href),
                }));
        }"""
    )

    # Body cell layout (header has an extra leading "Data pager" th, so row idx = header idx - 1):
    #   0=blank, 1=blank, 2=Spec#, 3=CR#, 4=Rev#, 5=CRCat, 6=ImpactedVer,
    #   7=TargetRel, 8=Title, 9=WG TDoc#, 10=WG status, 11=WG meeting, 12=WG source,
    #   13=TSG TDoc#, 14=TSG status, 15=TSG meeting, 16=TSG source, 17=NewVer,
    #   18=WorkItems, 19=Remarks
    IDX_SPEC, IDX_CR, IDX_REL, IDX_TITLE, IDX_WG, IDX_TSG = 2, 3, 7, 8, 9, 13

    out = []
    for r in rows:
        cells = r["cells"]
        links = r["tdocLinks"]
        if len(cells) <= IDX_TSG:
            continue
        wg_text = cells[IDX_WG].strip()
        tsg_text = cells[IDX_TSG].strip()
        wg_uid = next((extract_uid(l) for l in links if extract_uid(l) == wg_text), None)
        tsg_uid = next((extract_uid(l) for l in links if extract_uid(l) == tsg_text), None)
        if not wg_uid and not tsg_uid:
            continue
        out.append({
            "spec": cells[IDX_SPEC].strip(),
            "cr": cells[IDX_CR].strip(),
            "rel": cells[IDX_REL].strip(),
            "title": cells[IDX_TITLE].strip()[:80],
            "wg_uid": wg_uid,
            "tsg_uid": tsg_uid,
        })
    return out


def extract_uid(href: str) -> str | None:
    m = re.search(r"contributionUid=([^&]+)", href)
    return m.group(1) if m else None


def resolve_zip_url(uid: str, session: requests.Session) -> str | None:
    """Hit DownloadTDoc.aspx wrapper page and extract the inline redirect URL."""
    r = session.get(DOWNLOAD_URL.format(uid=uid), timeout=30)
    if r.status_code != 200:
        return None
    if "TDoc cannot be found" in r.text:
        return None
    m = REDIRECT_RE.search(r.text)
    return m.group(1) if m else None


def download_and_extract(file_url: str, output_dir: Path, uid: str, kind: str, keep_zip: bool, session: requests.Session) -> list[Path]:
    """Download the URL; if zip, extract .doc/.docx members; otherwise save raw. Returns extracted file paths."""
    output_dir.mkdir(parents=True, exist_ok=True)
    r = session.get(file_url, timeout=120)
    if r.status_code != 200:
        raise RuntimeError(f"HTTP {r.status_code} for {file_url}")
    data = r.content
    fname = Path(urlparse(file_url).path).name
    saved: list[Path] = []

    if fname.lower().endswith(".zip"):
        if keep_zip:
            zip_path = output_dir / f"{uid}_{kind}.zip"
            zip_path.write_bytes(data)
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            for member in zf.namelist():
                base = Path(member).name
                # Skip macOS resource forks and any AppleDouble metadata
                if base.startswith("._") or "__MACOSX/" in member:
                    continue
                low = base.lower()
                if low.endswith(".doc") or low.endswith(".docx"):
                    safe_name = f"{uid}_{kind}__{base}"
                    dest = output_dir / safe_name
                    with zf.open(member) as src, open(dest, "wb") as dst:
                        dst.write(src.read())
                    saved.append(dest)
    else:
        # already a .doc/.docx/.pdf or unknown — save as-is
        dest = output_dir / f"{uid}_{kind}__{fname}"
        dest.write_bytes(data)
        saved.append(dest)

    return saved


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--n", type=int, default=15, help="Number of random CRs to download (default: 15)")
    ap.add_argument("--specs", nargs="+", default=DEFAULT_SPECS, help="Spec numbers to search across (default: popular mix)")
    ap.add_argument("--kinds", nargs="+", default=["wg", "tsg"], choices=["wg", "tsg"], help="Which TDoc columns to download (default: both)")
    ap.add_argument("--output", default="input/change_requests", help="Output directory (default: input/change_requests)")
    ap.add_argument("--keep-zip", action="store_true", help="Also keep the original .zip file")
    ap.add_argument("--headed", action="store_true", help="Show browser window (debug)")
    ap.add_argument("--seed", type=int, default=None, help="Random seed for reproducible sampling")
    args = ap.parse_args()

    if args.seed is not None:
        random.seed(args.seed)

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    # ── Step 1: scrape CR list per spec via Playwright ──
    all_rows: list[dict] = []
    print(f"Querying CRs for specs: {args.specs}")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=not args.headed)
        ctx = browser.new_context(accept_downloads=False)
        page = ctx.new_page()
        page.goto(PORTAL_URL, wait_until="networkidle", timeout=90_000)
        time.sleep(6)  # SPA tab + grid bootstrap

        for spec in tqdm(args.specs, desc="Specs", unit="spec"):
            try:
                rows = collect_uids_for_spec(page, spec)
                tqdm.write(f"  {spec}: {len(rows)} CR rows")
                all_rows.extend(rows)
            except Exception as e:
                tqdm.write(f"  {spec}: ERROR — {e}")

        browser.close()

    if not all_rows:
        print("No CRs collected.", file=sys.stderr)
        sys.exit(1)

    # ── Step 2: random sample N CRs ──
    sample = random.sample(all_rows, min(args.n, len(all_rows)))
    print(f"\nSampled {len(sample)} CRs:")
    for r in sample:
        print(f"  {r['spec']:>8}  CR#{r['cr']:<5}  {r['rel']:<8}  wg={r['wg_uid'] or '-':<16}  tsg={r['tsg_uid'] or '-':<16}  {r['title']}")

    # ── Step 3: download + extract each kind ──
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (docx_parsing test downloader)"})

    total_files = 0
    failures = 0
    pbar = tqdm(sample, desc="Downloading", unit="CR")
    for r in pbar:
        for kind in args.kinds:
            uid = r.get(f"{kind}_uid")
            if not uid:
                continue
            pbar.set_postfix_str(f"{uid} ({kind})")
            try:
                zip_url = resolve_zip_url(uid, session)
                if not zip_url:
                    tqdm.write(f"  [skip] {uid} — no redirect URL (file missing on server)")
                    failures += 1
                    continue
                saved = download_and_extract(zip_url, output_dir, uid, kind, args.keep_zip, session)
                if not saved:
                    tqdm.write(f"  [warn] {uid} — zip had no .doc/.docx inside ({zip_url})")
                    failures += 1
                else:
                    for s in saved:
                        tqdm.write(f"  ✓ {s.name}")
                    total_files += len(saved)
            except Exception as e:
                tqdm.write(f"  [err]  {uid} ({kind}) — {e}")
                failures += 1
    pbar.close()

    print("\n── Run summary ─────────────────────────────")
    print(f"  CRs sampled:    {len(sample)}")
    print(f"  Files extracted: {total_files}")
    print(f"  Failures:       {failures}")
    print(f"  Output dir:     {output_dir.resolve()}")


if __name__ == "__main__":
    main()
