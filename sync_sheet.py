"""Pull the "Thunder Projects and Pilots" Excel sheet and write sheet.js for the page.

    SHEET_URL='<SharePoint share link>' python3 scripts/sync_sheet.py
    python3 scripts/sync_sheet.py path/to/local.xlsx

The sheet lists one row per deployment, under heading rows (text in column A only):
  a section heading such as "Thunder", "Global IT" or "DC" starts a section, in page order;
  "Projects" and "Pilots" headings split a section into projects and pilots.
Columns are found by their header, so they can be reordered:
  Project Name (required) | deployments | Hyperlink | Status | Server | Website | Note
sheet.js is only rewritten when the data changes, so its timestamp is the last sheet edit.
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


def read_sheets(xlsx):
    """Every worksheet as {tab name: [(row number, {column letter: (text, hyperlink)})]}, in tab order."""
    z = zipfile.ZipFile(BytesIO(xlsx))
    names = z.namelist()
    shared = []
    if "xl/sharedStrings.xml" in names:
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS):
            shared.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
    targets = {rel.get("Id"): rel.get("Target") for rel in
               ET.fromstring(z.read("xl/_rels/workbook.xml.rels")).findall("pr:Relationship", NS)}
    sheets = {}
    for tab in ET.fromstring(z.read("xl/workbook.xml")).findall("m:sheets/m:sheet", NS):
        path = "xl/" + targets[tab.get(RID)].lstrip("/").removeprefix("xl/")
        rels_path = path.replace("worksheets/", "worksheets/_rels/") + ".rels"
        rels = {}
        if rels_path in names:
            for rel in ET.fromstring(z.read(rels_path)).findall("pr:Relationship", NS):
                rels[rel.get("Id")] = rel.get("Target")
        sheet = ET.fromstring(z.read(path))
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
        sheets[tab.get("name")] = rows
    return sheets


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


def header_columns(texts, columns):
    cols = {}
    for col, text in texts.items():
        field = next((f for key, f in columns.items() if text.lower().startswith(key)), None)
        if field and field not in cols.values():
            cols[col] = field
    return cols


def parse(xlsx):
    """(deployment rows, sections in sheet order) from the first tab with a "Project Name" header."""
    for cells_by_row in read_sheets(xlsx).values():
        rows, sections = parse_deployments(cells_by_row)
        if rows:
            return rows, sections
    return [], []


def parse_deployments(cells_by_row):
    section, group, cols, out, sections = "", "projects", {}, [], []
    for _, cells in cells_by_row:
        texts = {col: text for col, (text, _) in cells.items()}
        if not cols:
            # Everything above the header row (the one with "Project Name") is ignored
            cols = header_columns(texts, COLUMNS)
            if "name" not in cols.values():
                cols = {}
            continue
        row = {field: texts.get(col, "") for col, field in cols.items()}
        if not row.get("name"):
            # Heading rows only fill column A: "Projects"/"Pilots" split a section, anything else starts one
            a = texts.get("A", "")
            if a and not a.isdigit():
                if "pilot" in a.lower():
                    group = "pilots"
                elif "project" in a.lower():
                    group = "projects"
                else:
                    section, group = a, "projects"
                    if a.lower() not in (x["name"].lower() for x in sections):
                        sections.append({"name": a})
            continue
        url_col = next((c for c, f in cols.items() if f == "url"), None)
        link = cells.get(url_col, ("", ""))[1] if url_col else ""
        text = row.get("url", "")
        row["url"] = link or (text if re.match(r"https?://", text) else "")
        row["env"] = env_of(row.get("env", ""), group)
        # An old-style Section column, if still there, wins over the heading
        row["section"] = row.get("section") or section
        out.append({"group": group, **{k: v for k, v in row.items() if v}})
    return out, sections


def write(rows, sections):
    """Rewrite sheet.js if the rows or sections changed. Returns True when it did."""
    old = None
    if OUT.exists():
        m = re.search(r"=\s*(\{.*\});?\s*$", OUT.read_text(), re.S)
        old = json.loads(m.group(1)) if m else None
    if old and old.get("rows") == rows and old.get("sections", []) == sections:
        print(f"sheet.js already up to date ({len(rows)} rows, {len(sections)} sections)")
        return False
    data = {"updated": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"), "sections": sections, "rows": rows}
    OUT.write_text(
        "// Generated by scripts/sync_sheet.py from the \"Thunder Projects and Pilots\" sheet. Do not edit;\n"
        "// change the sheet instead. Server specs live in deployments.js.\n"
        "window.THUNDER_SHEET = " + json.dumps(data, indent=1, ensure_ascii=False) + ";\n")
    print(f"Wrote sheet.js ({len(rows)} rows, {len(sections)} sections)")
    return True


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("SHEET_URL", "")
    if not src:
        sys.exit("Pass a local .xlsx path or set SHEET_URL to the sheet's share link.")
    xlsx = download(src) if re.match(r"https?://", src) else Path(src).read_bytes()
    rows, sections = parse(xlsx)
    if not rows:
        sys.exit("No rows found in the sheet; leaving sheet.js unchanged.")
    write(rows, sections)


if __name__ == "__main__":
    main()
