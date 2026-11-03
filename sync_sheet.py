"""Pull the "Thunder Projects and Pilots" Excel sheet and write sheet.js for the page.

    SHEET_URL='<SharePoint share link>' python3 scripts/sync_sheet.py
    python3 scripts/sync_sheet.py path/to/local.xlsx

The sheet is the list of deployments: one row per deployment, under "Projects" and
"Pilots" heading rows. Columns are found by their header, so they can be reordered:
  Project Name (required) | deployments | Hyperlink | Section | Status | Server | Website | Note
sheet.js is only rewritten when the rows change, so its timestamp is the last sheet edit.
Uses the standard library only (runs as-is in GitHub Actions).
"""
import http.cookiejar
import json
import os
import re
import sys
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

HERE = Path(__file__).resolve().parent
# Lives in scripts/ in the repo, but next to index.html when uploaded by hand (see build.py)
ROOT = HERE if (HERE / "index.html").exists() else HERE.parent
OUT = ROOT / "sheet.js"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def download(share_url):
    # "Anyone with the link" shares set a guest cookie on first visit; the download needs it
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
    opener.addheaders = [("User-Agent", "Mozilla/5.0 thunder-deployments-sync")]
    opener.open(share_url, timeout=60).read()
    sep = "&" if "?" in share_url else "?"
    data = opener.open(share_url + sep + "download=1", timeout=60).read()
    if not data.startswith(b"PK"):
        sys.exit("Download did not return an .xlsx file. Is the share link still 'Anyone with the link'?")
    return data


def read_cells(xlsx):
    """First worksheet as [(row number, {column letter: (text, hyperlink)})]."""
    z = zipfile.ZipFile(BytesIO(xlsx))
    names = z.namelist()
    shared = []
    if "xl/sharedStrings.xml" in names:
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS):
            shared.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
    rels = {}
    if "xl/worksheets/_rels/sheet1.xml.rels" in names:
        for rel in ET.fromstring(z.read("xl/worksheets/_rels/sheet1.xml.rels")).findall("pr:Relationship", NS):
            rels[rel.get("Id")] = rel.get("Target")
    sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    links = {h.get("ref"): rels.get(h.get(RID), "") for h in sheet.findall("m:hyperlinks/m:hyperlink", NS)}
    rows = []
    for row in sheet.findall("m:sheetData/m:row", NS):
        cells = {}
        for c in row.findall("m:c", NS):
            ref = c.get("r")
            v = c.find("m:v", NS)
            if c.get("t") == "s" and v is not None:
                text = shared[int(v.text)]
            elif c.get("t") == "inlineStr":
                text = "".join(t.text or "" for t in c.iter("{%s}t" % NS["m"]))
            else:
                text = v.text if v is not None else ""
            text = " ".join((text or "").split())
            if text or links.get(ref):
                cells[re.match(r"[A-Z]+", ref).group()] = (text, links.get(ref, ""))
        rows.append((int(row.get("r")), cells))
    return rows


def env_of(raw, group):
    raw = raw.lower()
    if "dev" in raw:
        return "dev"
    if "prod" in raw:
        return "prod"
    if "pilot" in raw or "trial" in raw or "trail" in raw:
        return "trial"
    return "trial" if group == "pilots" else "prod"


# Header text (lowercased) -> output field; a header matches when it starts with the key
COLUMNS = {"project name": "name", "deployment": "env", "hyperlink": "url", "section": "section",
           "status": "status", "server": "server", "website": "website", "note": "note"}


def parse(xlsx):
    group, cols, out = "projects", {}, []
    for _, cells in read_cells(xlsx):
        texts = {col: text for col, (text, _) in cells.items()}
        if not cols:
            # Everything above the header row (the one with "Project Name") is ignored
            for col, text in texts.items():
                field = next((f for key, f in COLUMNS.items() if text.lower().startswith(key)), None)
                if field and field not in cols.values():
                    cols[col] = field
            if "name" not in cols.values():
                cols = {}
            continue
        row = {field: texts.get(col, "") for col, field in cols.items()}
        if not row.get("name"):
            # Heading rows ("Projects", "Pilots") only fill column A
            a = texts.get("A", "")
            if a and not a.isdigit():
                group = "pilots" if "pilot" in a.lower() else re.sub(r"[^a-z0-9]+", "-", a.lower()).strip("-")
            continue
        url_col = next((c for c, f in cols.items() if f == "url"), None)
        link = cells.get(url_col, ("", ""))[1] if url_col else ""
        text = row.get("url", "")
        row["url"] = link or (text if re.match(r"https?://", text) else "")
        row["env"] = env_of(row.get("env", ""), group)
        out.append({"group": group, **{k: v for k, v in row.items() if v}})
    return out


def write(rows):
    """Rewrite sheet.js if the rows changed. Returns True when it did."""
    old = None
    if OUT.exists():
        m = re.search(r"=\s*(\{.*\});?\s*$", OUT.read_text(), re.S)
        old = json.loads(m.group(1)) if m else None
    if old and old.get("rows") == rows:
        print(f"sheet.js already up to date ({len(rows)} rows)")
        return False
    data = {"updated": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"), "rows": rows}
    OUT.write_text(
        "// Generated by scripts/sync_sheet.py from the \"Thunder Projects and Pilots\" sheet. Do not edit;\n"
        "// change the sheet instead. Section names and server specs live in deployments.js.\n"
        "window.THUNDER_SHEET = " + json.dumps(data, indent=1, ensure_ascii=False) + ";\n")
    print(f"Wrote sheet.js ({len(rows)} rows)")
    return True


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("SHEET_URL", "")
    if not src:
        sys.exit("Pass a local .xlsx path or set SHEET_URL to the sheet's share link.")
    xlsx = download(src) if re.match(r"https?://", src) else Path(src).read_bytes()
    rows = parse(xlsx)
    if not rows:
        sys.exit("No rows found in the sheet; leaving sheet.js unchanged.")
    write(rows)


if __name__ == "__main__":
    main()
