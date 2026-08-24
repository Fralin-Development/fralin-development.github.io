import urllib.request
import http.cookiejar
import zipfile
import xml.etree.ElementTree as ET
import os
import subprocess

SHAREPOINT_URL = os.environ.get(
    "SHAREPOINT_URL",
    "https://myuva-my.sharepoint.com/:x:/g/personal/dbg5af_virginia_edu/IQBDrHqx3UG3QJppnFzTHoMGATxyrU3UWV3-Irso_3Nirt8?download=1"
)
TARGET_TAB_NAME = os.environ.get("TARGET_TAB_NAME", "Donation Screen")
OUTPUT_CSV = os.environ.get("OUTPUT_CSV", "donors.csv")
TEMP_XLSX = "/tmp/downloaded_sheet.xlsx"
COOKIE_JAR = "/tmp/sp_cookies.txt"

def download_workbook():
    print("Downloading workbook from SharePoint via curl with cookie jar...")
    cmd = [
        "curl", "-s", "-L",
        "-c", COOKIE_JAR,
        "-b", COOKIE_JAR,
        SHAREPOINT_URL,
        "-o", TEMP_XLSX
    ]
    res = subprocess.run(cmd, capture_output=True)
    if res.returncode != 0:
        raise RuntimeError(f"curl failed: {res.stderr.decode('utf-8')}")

def sync():
    download_workbook()

    print("Extracting workbook data...")
    with zipfile.ZipFile(TEMP_XLSX, "r") as z:
        wb_xml = z.read("xl/workbook.xml")
        wb_root = ET.fromstring(wb_xml)
        sheets = wb_root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet")
        
        target_sheet_rId = None
        for s in sheets:
            name = s.attrib.get("name", "")
            if name.lower().strip() == TARGET_TAB_NAME.lower().strip() or TARGET_TAB_NAME.lower() in name.lower():
                target_sheet_rId = s.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                print(f"Matched sheet: '{name}' (rId: {target_sheet_rId})")
                break

        rels_xml = z.read("xl/_rels/workbook.xml.rels")
        rels_root = ET.fromstring(rels_xml)
        sheet_filename = "worksheets/sheet4.xml"
        for rel in rels_root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            if rel.attrib.get("Id") == target_sheet_rId:
                sheet_filename = rel.attrib.get("Target")
                break

        sheet_path = f"xl/{sheet_filename}" if not sheet_filename.startswith("xl/") else sheet_filename
        print(f"Reading worksheet from {sheet_path}...")

        shared_strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss_xml = z.read("xl/sharedStrings.xml")
            ss_root = ET.fromstring(ss_xml)
            for si in ss_root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"):
                text = "".join([t.text or "" for t in si.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")])
                shared_strings.append(text)

        sh_xml = z.read(sheet_path)
        sh_root = ET.fromstring(sh_xml)
        rows = sh_root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row")
        
        donor_names = []
        for r in rows:
            row_cells = []
            for c in r.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
                v = c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
                t = c.attrib.get("t")
                if v is not None:
                    val = shared_strings[int(v.text)] if t == "s" else v.text
                    val = val.strip()
                    if val:
                        row_cells.append(val)
            if row_cells and row_cells[0]:
                donor_names.append(row_cells[0])

    if not donor_names:
        print("Warning: No donors found in target sheet.")
        return

    print(f"Found {len(donor_names)} donors. Writing to {OUTPUT_CSV}...")
    with open(OUTPUT_CSV, "w", encoding="utf-8") as f:
        f.write("Donor Name\n")
        for name in donor_names:
            clean_name = name.replace('"', '""')
            f.write(f'"{clean_name}"\n')

    print("Sync complete.")

if __name__ == "__main__":
    sync()
