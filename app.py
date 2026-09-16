"""Automates the Accenture 'Closed with Insuff' export from the AuthBridge MIS query browser.

Logs into https://mis.authbridge.com/export_query/, selects Host/Database/Data Time
Slab/Query, fills the date range and client name, clicks Export, waits for the
downloaded .zip, extracts the CSV into this folder, and removes the zip.
"""
import math
import os
import time
import zipfile
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

load_dotenv()

LOGIN_URL = "https://mis.authbridge.com/export_query/login.php"
USERNAME = os.environ["MIS_USERNAME"]
PASSWORD = os.environ["MIS_PASSWORD"]

HOST = "Bridge Live"
DATABASE = "Bridge Live"
DATA_TIME_SLAB = "Morning Data - Daily Dump"
QUERY_NAME = "Yesterday Done checks with other details as per client name - Bridge"
CLIENT_NAME = "Accenture Solutions Private Limited"

ADVANCE_DATA_TIME_SLAB = "case query - Bridge"
ADVANCE_QUERY_NAME = "Advance Tracker"

ANTECEDENT_DATA_TIME_SLAB = "checks query - Bridge"
ANTECEDENT_QUERY_NAME = "Antecedent Details"

DOWNLOAD_DIR = Path(__file__).resolve().parent

CHECK_STATUS_FILTER = "Closed with Insufficiency"

# Antecedent Details rows whose stated_data feeds the HRT column.
HRT_FIELD_NAMES = ["Task ID", "Old Task Number", "New Task Number"]

# Default mapping file names looked for next to app.py when no mapping is
# supplied explicitly (e.g. from the Streamlit uploader).
MAPPING_FILE_NAMES = ("mapping.xlsx", "mapping.xls.xlsx")

# Advance Tracker's Check_unique_name -> Component label for column D.
# Used only if no mapping.xlsx (Check_unique_name / Mapped Name columns) is found.
DEFAULT_COMPONENT_MAP = {
    "Date of Birth Verification": "Legal Age to Work",
    "Education Verification W": "Education",
    "India Court Record Database Check": "Criminal Court Check",
    "Previous Employment Verification": "Previous Employment",
    "Global Database Check": "Criminal Database Check",
    "National Identity Check": "Identity Proof Check",
    "Current Employment Company Check": "Current Company",
    "Previous Employment Company Check": "Previous Company",
    "Current Employment Verification": "Current Employment",
    "Professional Reference Check": "Professional Reference Check",
    "Instant Previous Employment Check": "Previous Employment - Tier 3",
    "Instant Current Employment Check": "Current Employment - Tier 3",
    "Education Institution Validation": "Education Institution Validation",
    "India Court Record Check through Law Firm": "India Court Record Check through Law Firm",
}

# (source CSV column, output column header). None = leave blank (filled in by a
# future logic update). "__vendor__" = static "Authbridge" value. The
# "__advance_*__" markers are looked up from the Advance Tracker export and
# "__antecedent_hrt__" from the Antecedent Details export, both matched by
# Case_Check_id / case_check_id.
COLUMN_MAP = [
    ("case_flex_field1", "CID"),
    ("case_flex_field6", "HRC"),
    ("__antecedent_hrt__", "HRT"),
    ("__advance_component__", "Component"),
    ("insuff_remarks", "Insufficiency Details"),
    ("__advance_verification_source__", "University/College/Employment/Company - Name"),
    ("case_received_date", "Date of Initiation"),
    ("__advance_first_insuff_date__", "Insuff Raised date"),
    ("Insuff_fulfill_date", "Insuff Closed date"),
    ("__vendor__", "Vendor Name"),
    ("case_check_id", "case_check_id"),
    ("case_ars_no", "case_ars_no"),
]

# Date columns rendered as dd-mmm-yy (e.g. 09-Jul-26). "Insuff Raised date"
# keeps its full timestamp value underneath the cell - only the display is
# date-only; the raw date and time are still there if you click into it.
DATE_DISPLAY_COLUMNS = {"Date of Initiation", "Insuff Raised date", "Insuff Closed date"}
DATE_NUMBER_FORMAT = "dd-mmm-yy"

# Columns wide enough to hold free-text remarks; wrapped instead of truncated.
WRAP_COLUMNS = {"Insufficiency Details", "University/College/Employment/Company - Name"}

COLUMN_WIDTHS = {
    "CID": 16,
    "HRC": 14,
    "HRT": 16,
    "Component": 24,
    "Insufficiency Details": 55,
    "University/College/Employment/Company - Name": 38,
    "Date of Initiation": 16,
    "Insuff Raised date": 16,
    "Insuff Closed date": 16,
    "Vendor Name": 13,
    "case_check_id": 15,
    "case_ars_no": 14,
}

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
BAND_FILL = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
GRID_SIDE = Side(style="thin", color="D9D9D9")
GRID_BORDER = Border(left=GRID_SIDE, right=GRID_SIDE, top=GRID_SIDE, bottom=GRID_SIDE)


def compute_date_range():
    """Yesterday's date range, except on Monday where it spans Fri-Sun."""
    today = date.today()
    if today.weekday() == 0:  # Monday
        from_date = today - timedelta(days=3)  # Friday
        to_date = today - timedelta(days=1)  # Sunday
    else:
        from_date = today - timedelta(days=1)
        to_date = from_date
    return f"{from_date} 00:00:00", f"{to_date} 23:59:59"


def build_driver(download_dir):
    options = Options()
    if os.environ.get("HEADLESS", "true").lower() != "false":
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1600,1000")
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": str(download_dir),
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True,
        },
    )
    driver = webdriver.Chrome(options=options)
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": str(download_dir)},
    )
    return driver


def select_dropdown(driver, name, text, timeout=20, attempts=3):
    """Select an option by visible text, retrying on stale elements.

    The site rebuilds each dependent <select> via AJAX after its parent
    changes, so the node backing `name` can be swapped out between checking
    its options and clicking one; retry the whole select on that race.
    """
    locator = (By.NAME, name)
    for attempt in range(attempts):
        try:
            WebDriverWait(driver, timeout, ignored_exceptions=(StaleElementReferenceException,)).until(
                lambda d: len(Select(d.find_element(*locator)).options) > 1
            )
            Select(driver.find_element(*locator)).select_by_visible_text(text)
            return
        except StaleElementReferenceException:
            if attempt == attempts - 1:
                raise
            time.sleep(0.5)


def _retry_stale(action, attempts=3, delay=0.5):
    """Run `action`, retrying it from scratch if the AJAX-rebuilt DOM makes
    an element go stale mid-interaction."""
    for attempt in range(attempts):
        try:
            return action()
        except StaleElementReferenceException:
            if attempt == attempts - 1:
                raise
            time.sleep(delay)


def login(driver):
    driver.get(LOGIN_URL)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "username")))
    driver.find_element(By.NAME, "username").send_keys(USERNAME)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    driver.find_element(By.NAME, "login").click()
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "hostname")))


def fill_query_form(driver, from_date, to_date):
    select_dropdown(driver, "hostname", HOST)
    select_dropdown(driver, "database", DATABASE)
    select_dropdown(driver, "access_time", DATA_TIME_SLAB)
    select_dropdown(driver, "csv_query", QUERY_NAME)

    WebDriverWait(driver, 20, ignored_exceptions=(StaleElementReferenceException,)).until(
        EC.presence_of_element_located((By.NAME, "date1"))
    )

    def fill_and_submit():
        date1 = driver.find_element(By.NAME, "date1")
        date2 = driver.find_element(By.NAME, "date2")
        client1 = driver.find_element(By.NAME, "client1")

        date1.clear()
        date1.send_keys(from_date)
        date2.clear()
        date2.send_keys(to_date)
        client1.clear()
        client1.send_keys(CLIENT_NAME)

        driver.find_element(By.ID, "run_query").click()

    _retry_stale(fill_and_submit)


def fill_check_id_query_form(driver, data_time_slab, query_name, value):
    """Fill the shared 'check_id1' text field used by both the Advance
    Tracker (ARS No*) and Antecedent Details (Check Id:*) queries."""
    select_dropdown(driver, "access_time", data_time_slab)
    select_dropdown(driver, "csv_query", query_name)

    WebDriverWait(driver, 20, ignored_exceptions=(StaleElementReferenceException,)).until(
        EC.presence_of_element_located((By.NAME, "check_id1"))
    )

    def fill_and_submit():
        field = driver.find_element(By.NAME, "check_id1")
        field.clear()
        field.send_keys(value)
        driver.find_element(By.ID, "run_query").click()

    _retry_stale(fill_and_submit)


def wait_for_download(download_dir, seen_before, timeout=120):
    """Wait for a new *.zip to appear. Chrome only uses the final `.zip`
    name once a download is complete (in-progress files are suffixed
    `.crdownload`), so a new zip's mere presence proves it's done - do not
    also gate on the absence of *any* .crdownload, since unrelated browser
    downloads (e.g. Chrome's own background component fetches) can leave
    stray .crdownload files that never resolve and would block forever."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        new_zips = [p for p in download_dir.glob("*.zip") if p not in seen_before]
        if new_zips:
            new_zips.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            return new_zips[0]
        time.sleep(1)
    raise TimeoutError("Timed out waiting for the export file to download.")


def extract_and_cleanup(zip_path):
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(zip_path.parent)
        names = zf.namelist()
    zip_path.unlink()
    return [zip_path.parent / name for name in names]


def load_closed_with_insuff(csv_path):
    df = pd.read_csv(csv_path, encoding="latin-1")
    return df[df["check_status"] == CHECK_STATUS_FILTER]


def build_hrt_lookup(antecedent_df):
    """case_check_id -> sorted, de-duplicated Task/Old/New Task Number
    stated_data values (comma-joined) for the HRT column."""
    sub = antecedent_df[antecedent_df["field_name"].isin(HRT_FIELD_NAMES)].copy()
    sub["stated_data"] = sub["stated_data"].astype(str).str.strip()
    sub = sub[~sub["stated_data"].isin(["", "nan", "Not Mentioned"])]
    sub["case_check_id"] = sub["case_check_id"].astype(str)
    return sub.groupby("case_check_id")["stated_data"].agg(lambda values: ", ".join(sorted(set(values))))


def load_component_map(source):
    """Read a Check_unique_name -> Mapped Name lookup from an Excel file
    (columns 'Check_unique_name' and 'Mapped Name'). `source` may be a path
    or a file-like object, e.g. a Streamlit file upload."""
    mapping_df = pd.read_excel(source)
    return dict(zip(mapping_df["Check_unique_name"], mapping_df["Mapped Name"]))


def find_default_mapping_file(directory):
    for name in MAPPING_FILE_NAMES:
        candidate = directory / name
        if candidate.exists():
            return candidate
    return None


def estimate_row_height(row, headers):
    """Approximate the row height (points) needed for wrapped text columns,
    since Excel won't auto-fit row height for content it hasn't rendered yet."""
    max_lines = 1
    for header, value in zip(headers, row):
        if header in WRAP_COLUMNS and value:
            chars_per_line = max(int(COLUMN_WIDTHS.get(header, 30) * 1.7), 10)
            max_lines = max(max_lines, math.ceil(len(str(value)) / chars_per_line))
    return min(max_lines * 15, 250)


def write_formatted_report(out, xlsx_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Closed with Insuff"

    headers = list(out.columns)
    ws.append(headers)
    ws.row_dimensions[1].height = 28
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = GRID_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_idx, row in enumerate(out.itertuples(index=False), start=2):
        for col_idx, (header, value) in enumerate(zip(headers, row), start=1):
            if pd.isna(value) or value == "":
                value = None
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = GRID_BORDER
            if row_idx % 2 == 0:
                cell.fill = BAND_FILL
            if header in DATE_DISPLAY_COLUMNS and value is not None:
                cell.number_format = DATE_NUMBER_FORMAT
            if header in WRAP_COLUMNS:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row_idx].height = estimate_row_height(row, headers)

    for col_idx, header in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = COLUMN_WIDTHS.get(header, 16)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(xlsx_path)


def build_insuff_report(closed, advance_df, antecedent_df, out_dir, component_map=None):
    component_map = component_map or DEFAULT_COMPONENT_MAP
    check_ids = closed["case_check_id"].astype(str)

    if advance_df is not None and not advance_df.empty:
        advance_lookup = advance_df.drop_duplicates("Case_Check_id").set_index("Case_Check_id")
        advance_lookup.index = advance_lookup.index.astype(str)
        first_insuff_date = check_ids.map(advance_lookup["First_Insuff_Date"])
        verification_source = check_ids.map(advance_lookup["verification_source"])

        check_unique_name = check_ids.map(advance_lookup["Check_unique_name"])
        component = check_unique_name.map(component_map)
        unmapped = sorted(check_unique_name[check_unique_name.notna() & component.isna()].unique())
        if unmapped:
            print(f"Warning: no Component mapping for Check_unique_name values: {unmapped}")
    else:
        first_insuff_date = ""
        verification_source = ""
        component = ""

    if antecedent_df is not None and not antecedent_df.empty:
        hrt = check_ids.map(build_hrt_lookup(antecedent_df))
    else:
        hrt = ""

    out = pd.DataFrame()
    for source, header in COLUMN_MAP:
        if source is None:
            out[header] = ""
        elif source == "__vendor__":
            out[header] = "Authbridge"
        elif source == "__advance_first_insuff_date__":
            out[header] = first_insuff_date
        elif source == "__advance_verification_source__":
            out[header] = verification_source
        elif source == "__antecedent_hrt__":
            out[header] = hrt
        elif source == "__advance_component__":
            out[header] = component
        else:
            out[header] = closed[source]

    out["Date of Initiation"] = pd.to_datetime(out["Date of Initiation"], errors="coerce")
    out["Insuff Raised date"] = pd.to_datetime(out["Insuff Raised date"], errors="coerce")
    out["Insuff Closed date"] = pd.to_datetime(out["Insuff Closed date"], errors="coerce")

    xlsx_path = out_dir / f"Accenture_Closed_with_Insuff_{date.today()}.xlsx"
    write_formatted_report(out, xlsx_path)
    return xlsx_path


def build_ars_string(closed):
    """Comma-separated, single-quoted ARS numbers (no trailing comma) for the
    MIS 'ARS No*' field, e.g. '5269-080960','5269-075588'."""
    unique_ars = list(dict.fromkeys(closed["case_ars_no"]))
    return ",".join(f"'{ars}'" for ars in unique_ars)


def build_check_id_string(closed):
    """Comma-separated case_check_id values (no trailing comma, no quotes)
    for the MIS 'Check Id:*' field, e.g. 2159938764,2158421704."""
    unique_ids = list(dict.fromkeys(closed["case_check_id"].astype(str)))
    return ",".join(unique_ids)


def run_pipeline(component_map=None, log=print):
    """Run the full export -> filter -> report pipeline and return a dict of
    the produced file paths (plus row counts) keyed by stage name."""
    from_date, to_date = compute_date_range()
    log(f"Requesting export for {from_date} -> {to_date}")

    results = {"from_date": from_date, "to_date": to_date}
    driver = build_driver(DOWNLOAD_DIR)
    try:
        seen_before = set(DOWNLOAD_DIR.glob("*.zip"))
        login(driver)
        fill_query_form(driver, from_date, to_date)
        zip_path = wait_for_download(DOWNLOAD_DIR, seen_before)
        extracted = extract_and_cleanup(zip_path)
        for path in extracted:
            log(f"Saved: {path}")
        results["daily_csv"] = extracted[0]

        closed = load_closed_with_insuff(extracted[0])
        results["closed_rows"] = len(closed)
        ars_string = build_ars_string(closed)

        advance_df = None
        if ars_string:
            seen_before = set(DOWNLOAD_DIR.glob("*.zip"))
            fill_check_id_query_form(driver, ADVANCE_DATA_TIME_SLAB, ADVANCE_QUERY_NAME, ars_string)
            zip_path = wait_for_download(DOWNLOAD_DIR, seen_before)
            advance_paths = extract_and_cleanup(zip_path)
            for path in advance_paths:
                log(f"Saved: {path}")
            advance_df = pd.read_csv(advance_paths[0], encoding="latin-1")
            results["advance_csv"] = advance_paths[0]
        else:
            log("No 'Closed with Insufficiency' rows today; skipping Advance Tracker export.")

        check_id_string = build_check_id_string(closed)
        antecedent_df = None
        if check_id_string:
            seen_before = set(DOWNLOAD_DIR.glob("*.zip"))
            fill_check_id_query_form(
                driver, ANTECEDENT_DATA_TIME_SLAB, ANTECEDENT_QUERY_NAME, check_id_string
            )
            zip_path = wait_for_download(DOWNLOAD_DIR, seen_before)
            antecedent_paths = extract_and_cleanup(zip_path)
            for path in antecedent_paths:
                log(f"Saved: {path}")
            antecedent_df = pd.read_csv(antecedent_paths[0], encoding="latin-1")
            results["antecedent_csv"] = antecedent_paths[0]
        else:
            log("No 'Closed with Insufficiency' rows today; skipping Antecedent Details export.")

        xlsx_path = build_insuff_report(closed, advance_df, antecedent_df, DOWNLOAD_DIR, component_map)
        log(f"Saved: {xlsx_path}")
        results["report_xlsx"] = xlsx_path
    finally:
        driver.quit()

    return results


def main():
    mapping_path = find_default_mapping_file(DOWNLOAD_DIR)
    component_map = load_component_map(mapping_path) if mapping_path else DEFAULT_COMPONENT_MAP
    run_pipeline(component_map)


if __name__ == "__main__":
    main()
