#!/usr/bin/env python3
"""
Fetches each "overview_video"/"video" lesson page in the Plan sheet and finds
the first YouTube video embedded in the main lesson content (sidebar, footer,
nav, comments and "related posts" are ignored). The watch URL is written into
a new "Video URL" column in data/IELTS_30_Day_Plan.xlsx -- the original URL
column is left untouched.

The spreadsheet stays the source of truth. Re-run any time after adding rows:
    python3 scripts/fetch-videos.py

Prints two lists at the end: pages with no video found, and pages with more
than one video found in the main content (first one was used -- check these).
"""
import re
import sys
import time
from copy import copy
from pathlib import Path
from urllib.parse import urlparse

try:
    import requests
    from bs4 import BeautifulSoup
    import openpyxl
    from openpyxl.utils import get_column_letter
except ImportError as e:
    sys.exit(f"Missing dependency: {e}. Run: pip3 install requests beautifulsoup4 lxml openpyxl")

ROOT = Path(__file__).resolve().parent.parent
XLSX_PATH = ROOT / "data" / "IELTS_30_Day_Plan.xlsx"

VIDEO_COL_NAME = "Video URL"
SCRAPE_TYPES = ("overview_video", "video")

EMBED_RE = re.compile(r'(?:youtube(?:-nocookie)?\.com/embed/|youtu\.be/)([A-Za-z0-9_-]{6,})')
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
}

# Removed from the whole document before searching, regardless of whether a
# main-content container is found, so sidebar/footer/nav/comment videos never count.
NOISE_SELECTORS = [
    "aside", "footer", "nav", "header", "script", "style", "noscript",
    "[class*=sidebar]", "[id*=sidebar]", "[class*=widget-area]",
    "[class*=comment]", "[id*=comment]",
    "[class*=related]", "[id*=related]",
]
# Tried in order; first one that matches something is treated as "the main lesson content".
MAIN_SELECTORS = [".entry-content", "article .entry-content", "main .entry-content", "article", "main", "#content"]


def find_main_video_ids(html):
    soup = BeautifulSoup(html, "lxml")

    # Find the main-content container BEFORE stripping anything: a layout
    # class like Genesis's "content-sidebar" on <body> means "this page HAS a
    # sidebar", not "this element IS a sidebar" -- decomposing by substring
    # match against page-wide containers (body/html) would wipe everything.
    scope = None
    for sel in MAIN_SELECTORS:
        scope = soup.select_one(sel)
        if scope:
            break
    if scope is None:
        scope = soup

    # Only now remove noise, and only from within the chosen scope, so a
    # layout class on an ancestor (body, wrapper divs) can never match.
    for sel in NOISE_SELECTORS:
        for tag in scope.select(sel):
            if tag.name in ("html", "body") or tag is scope:
                continue  # never nuke the whole page/scope over a layout class
            tag.decompose()

    ids = []
    seen = set()

    def add(val):
        if not val:
            return
        m = EMBED_RE.search(val)
        if m and m.group(1) not in seen:
            seen.add(m.group(1))
            ids.append(m.group(1))

    for iframe in scope.find_all("iframe"):
        add(iframe.get("src"))
        add(iframe.get("data-src"))
        add(iframe.get("data-lazy-src"))
    if not ids:
        for tag in scope.find_all(["a", "div"]):
            add(tag.get("href"))
            add(tag.get("data-src"))
            add(tag.get("data-lazy-src"))
    return ids


def fetch(url, tries=3):
    last_exc = None
    for attempt in range(tries):
        try:
            r = requests.get(url, headers=HEADERS, timeout=25)
            r.raise_for_status()
            return r.text
        except Exception as e:
            last_exc = e
            if attempt < tries - 1:
                time.sleep(1.5)
    raise last_exc


def main():
    if not XLSX_PATH.exists():
        sys.exit(f"Spreadsheet not found: {XLSX_PATH}")

    wb = openpyxl.load_workbook(XLSX_PATH)
    ws = wb["Plan"]

    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    idx = {h: i for i, h in enumerate(header) if h}
    for col in ("Day", "Item Type", "Title", "URL"):
        if col not in idx:
            sys.exit(f"Plan sheet is missing expected column: {col}")

    url_col = idx["URL"] + 1
    if VIDEO_COL_NAME in idx:
        video_col = idx[VIDEO_COL_NAME] + 1
    else:
        video_col = ws.max_column + 1
        header_cell = ws.cell(row=1, column=video_col, value=VIDEO_COL_NAME)
        src_header = ws.cell(row=1, column=url_col)
        header_cell.font = copy(src_header.font)
        header_cell.fill = copy(src_header.fill)
        header_cell.border = copy(src_header.border)
        header_cell.alignment = copy(src_header.alignment)
        ws.column_dimensions[get_column_letter(video_col)].width = 46

    no_video, multi_video = [], []
    updated = skipped = 0

    for r in range(2, ws.max_row + 1):
        item_type = ws.cell(row=r, column=idx["Item Type"] + 1).value
        if item_type not in SCRAPE_TYPES:
            continue
        url_cell = ws.cell(row=r, column=url_col)
        url = url_cell.value
        day = ws.cell(row=r, column=idx["Day"] + 1).value
        title = ws.cell(row=r, column=idx["Title"] + 1).value or ""
        vcell = ws.cell(row=r, column=video_col)

        if not url or str(url).strip() in ("", "—"):
            skipped += 1
            continue
        if urlparse(str(url)).scheme not in ("http", "https"):
            no_video.append((day, url, title, "not a fetchable URL"))
            continue

        print(f"Day {day}: {url}")
        try:
            html = fetch(url)
        except Exception as e:
            print(f"  ERROR fetching: {e}", file=sys.stderr)
            no_video.append((day, url, title, f"fetch error: {e}"))
            continue

        ids = find_main_video_ids(html)
        if not ids:
            no_video.append((day, url, title, "no video found in main content"))
            vcell.value = None
        else:
            if len(ids) > 1:
                multi_video.append((day, url, title, ids))
            watch_url = f"https://www.youtube.com/watch?v={ids[0]}"
            vcell.value = watch_url
            vcell.hyperlink = watch_url
            vcell.font = copy(url_cell.font)
            vcell.alignment = copy(url_cell.alignment)
            updated += 1
        time.sleep(0.6)  # be polite to the source site

    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    wb.save(XLSX_PATH)

    print(f"\nUpdated {updated} row(s) with a Video URL. Skipped {skipped} row(s) with no page URL.")

    print(f"\n=== Pages with NO video found in main content ({len(no_video)}) ===")
    for day, url, title, reason in no_video:
        print(f"  Day {day}: {title}\n    {url}\n    [{reason}]")

    print(f"\n=== Pages with MORE THAN ONE video in main content ({len(multi_video)}) ===")
    for day, url, title, ids in multi_video:
        print(f"  Day {day}: {title}\n    {url}\n    videos found: {ids}  -> used first: {ids[0]}")


if __name__ == "__main__":
    main()
