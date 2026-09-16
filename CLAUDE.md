# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Selenium-driven scraper + report builder that automates a daily "Accenture Closed with Insufficiency" pull from AuthBridge's internal MIS query browser (`https://mis.authbridge.com/export_query/`), then formats the result into a polished Excel deliverable. There are two entry points sharing the same pipeline:

- `app.py` — CLI, runs the full pipeline once (`python app.py`), meant for Task Scheduler / cron.
- `streamlit_app.py` — web UI over the same pipeline (`streamlit run streamlit_app.py`), adds an upload widget for the Component mapping file.

Everything lives in one module, `app.py`; `streamlit_app.py` imports functions from it rather than duplicating logic.

## Commands

```bash
pip install -r requirements.txt
python app.py                          # run the full pipeline once
streamlit run streamlit_app.py         # web UI version
```

No test suite, linter, or build step exists in this repo.

Credentials live in `.env` (git-ignored; `.env.example` shows the shape): `MIS_USERNAME`, `MIS_PASSWORD`, `HEADLESS` (default `true`; set `false` to watch Chrome run visibly for debugging).

## Architecture: the four-phase pipeline

`run_pipeline()` in `app.py` drives one Selenium/Chrome session through four sequential exports on the same MIS site, reusing the logged-in session throughout (the site doesn't navigate away on export — it just triggers a file download, so Host/Database selections persist across phases):

1. **Daily export** — Host=`Bridge Live`, Data Time Slab=`Morning Data - Daily Dump`, Query=`Yesterday Done checks with other details as per client name - Bridge`. Date range is `compute_date_range()`: yesterday, *except* on Monday where it spans Friday–Sunday (to cover the weekend gap). Filtered to `check_status == "Closed with Insufficiency"` (`load_closed_with_insuff`) — this `closed` DataFrame is the backbone the other three phases key off of.
2. **Advance Tracker export** — keyed by ARS numbers (`build_ars_string`) built from `closed`. Supplies `First_Insuff_Date`, `verification_source`, and `Check_unique_name` back onto each row via `case_check_id` = `Case_Check_id`.
3. **Antecedent Details export** — keyed by `case_check_id` values (`build_check_id_string`), *not* ARS numbers. Supplies the HRT column: filtered to `field_name` in `{Task ID, Old Task Number, New Task Number}`, deduped/sorted A→Z per check (`build_hrt_lookup`).
4. **Report build** (`build_insuff_report` → `write_formatted_report`) — joins all three exports into the final `Accenture_Closed_with_Insuff_<date>.xlsx`, formatted (see below).

If `closed` is empty for the day, phases 2–4 degrade gracefully (blank columns) rather than submitting empty queries to the site.

### A critical, non-obvious MIS-site quirk

The same underlying `check_id1` text input is reused by both phase 2 and phase 3, but **the expected value format differs by query**, discovered empirically (not documented anywhere on the site):
- Advance Tracker's "ARS No*" field needs **single-quoted, comma-separated** values: `'5269-080960','5269-075588'`. An unquoted list silently returns zero rows.
- Antecedent Details' "Check Id:*" field needs **plain comma-separated, unquoted** values: `2159938764,2158421704`.

Both are built with no trailing comma (`build_ars_string`, `build_check_id_string`).

### `COLUMN_MAP` is the single source of truth for the report shape

The output workbook's 12 columns (CID → case_ars_no) are defined by the `(source, header)` pairs in `COLUMN_MAP`. `source` is either a literal CSV column name, or one of the `"__advance_*__"` / `"__antecedent_hrt__"` markers resolved by lookup tables built earlier in `build_insuff_report`. When asked to change what feeds a column, edit this list plus the corresponding lookup — don't hardcode column letters elsewhere.

### Component mapping resolution order

`Check_unique_name → Component` (column D) is *not* hardcoded to be reliable long-term — it's meant to be maintained externally. Resolution order in `main()`: an uploaded file (Streamlit) → `mapping.xlsx` or `mapping.xls.xlsx` found next to `app.py` (`find_default_mapping_file`) → `DEFAULT_COMPONENT_MAP` fallback baked into the code. The mapping file format is fixed: columns `Check_unique_name` and `Mapped Name` (see `load_component_map`).

### Report formatting contract

`write_formatted_report` builds the workbook directly with `openpyxl` (not `df.to_excel`) for full style control: navy header row, thin gridlines, zebra striping, frozen header, autofilter, hand-tuned `COLUMN_WIDTHS` per column, and estimated wrapped-row heights (`estimate_row_height`) for the two free-text columns (`WRAP_COLUMNS`). The three date columns (`DATE_DISPLAY_COLUMNS`) are real `datetime` values formatted `dd-mmm-yy` for display — critically, "Insuff Raised date" keeps its full timestamp in the underlying cell value; only the *display* is date-only. Don't truncate that column to a date string — the time component is intentionally preserved.

### Selenium reliability quirks (don't remove these without understanding why)

- **Stale elements**: the site rebuilds each dependent `<select>`/text-input via AJAX whenever a parent field changes, so a node fetched a moment earlier can go stale mid-interaction. `select_dropdown` and `_retry_stale` retry the whole select-or-fill-and-submit operation as one atomic unit rather than retrying individual calls.
- **`wait_for_download`** only checks for the *appearance* of a new `.zip` file — it deliberately does **not** also gate on the absence of `.crdownload` files in the directory. Headless Chrome can leave unrelated stray `.crdownload` junk (e.g. its own background component/model fetches) that never resolves; gating on "no crdownloads anywhere" caused real timeouts in production even after the target file had finished downloading.
- The MIS site's login is a plain `<input type="submit" name="login">`, not a `<button>` — locate it by `By.NAME, "login"`, not by button text.

## Output files (written to the same directory as `app.py`)

- `Yesterday Done checks with other details as per client name - Bridge.csv` — raw daily export
- `Advance Tracker.csv`, `Antecedent Details.csv` — the two lookup exports
- `Accenture_Closed_with_Insuff_<date>.xlsx` — the final formatted deliverable

All are the direct extracted contents of the site's own zip downloads (zips are deleted after extraction); filenames come from the site, not chosen by this code.
