# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A report builder that queries the live Bridge MySQL database directly for the daily "Accenture Closed with Insufficiency" pull, then formats the result into a polished Excel deliverable. There are two entry points sharing the same pipeline:

- `app.py` — CLI, runs the full pipeline once (`python app.py`), meant for Task Scheduler / cron.
- `streamlit_app.py` — web UI over the same pipeline (`streamlit run streamlit_app.py`), adds an upload widget for the Component mapping file.

Pipeline logic lives in `app.py`, the three SQL statements in `queries.py`; `streamlit_app.py` imports functions from it rather than duplicating logic.

## Commands

```bash
pip install -r requirements.txt
python app.py                          # run the full pipeline once
streamlit run streamlit_app.py         # web UI version
```

No test suite, linter, or build step exists in this repo.

DB settings live in `.env` (git-ignored; `.env.example` shows the shape): `DB_HOST`, `DB_PORT`, `DB_USER`, `DB_PASSWORD`, `DB_NAME` (default `checkpoint_live`).

## Architecture: the four-phase pipeline

`run_pipeline()` in `app.py` opens one pymysql connection (`connect_db`) and runs three queries from `queries.py` (via `fetch_df` / `fetch_in_query`, which bind all values as parameters; literal `%` in the SQL is written `%%`):

1. **Daily query** (`DAILY_QUERY`) — date range is `compute_date_range()`: yesterday, *except* on Monday where it spans Friday–Sunday. Client is matched with `LIKE %CLIENT_NAME%`. Filtered to `check_status == "Closed with Insufficiency"` (`load_closed_with_insuff`) — this `closed` DataFrame is the backbone the other phases key off of.
2. **Advance Tracker query** (`ADVANCE_QUERY`) — keyed by ARS numbers (`build_ars_list`). Supplies `First_Insuff_Date`, `verification_source`, and `Check_unique_name` back onto each row via `case_check_id` = `Case_Check_id`. It uses the `@insuffdate` session variable, so it must run on the same connection.
3. **Antecedent Details query** (`ANTECEDENT_QUERY`) — keyed by `case_check_id` values (`build_check_id_list`). Supplies the HRT column: filtered to `field_name` in `{Task ID, Old Task Number, New Task Number}`, deduped/sorted A→Z per check (`build_hrt_lookup`).
4. **Report build** (`build_insuff_report` → `write_formatted_report`) — joins all three results into the final `Accenture_Closed_with_Insuff_<date>.xlsx`, formatted (see below).

If `closed` is empty for the day, phases 2–4 degrade gracefully (blank columns) rather than running empty `IN ()` queries. `fetch_df` renames returned columns case-insensitively to the spellings listed in `DAILY_COLUMNS` / `ADVANCE_COLUMNS` / `ANTECEDENT_COLUMNS`, since MySQL may report table-defined casing.

### `COLUMN_MAP` is the single source of truth for the report shape

The output workbook's 12 columns (CID → case_ars_no) are defined by the `(source, header)` pairs in `COLUMN_MAP`. `source` is either a literal CSV column name, or one of the `"__advance_*__"` / `"__antecedent_hrt__"` markers resolved by lookup tables built earlier in `build_insuff_report`. When asked to change what feeds a column, edit this list plus the corresponding lookup — don't hardcode column letters elsewhere.

### Component mapping resolution order

`Check_unique_name → Component` (column D) is *not* hardcoded to be reliable long-term — it's meant to be maintained externally. Resolution order in `main()`: an uploaded file (Streamlit) → `mapping.xlsx` or `mapping.xls.xlsx` found next to `app.py` (`find_default_mapping_file`) → `DEFAULT_COMPONENT_MAP` fallback baked into the code. The mapping file format is fixed: columns `Check_unique_name` and `Mapped Name` (see `load_component_map`).

### Report formatting contract

`write_formatted_report` builds the workbook directly with `openpyxl` (not `df.to_excel`) for full style control: navy header row, thin gridlines, zebra striping, frozen header, autofilter, hand-tuned `COLUMN_WIDTHS` per column, and estimated wrapped-row heights (`estimate_row_height`) for the two free-text columns (`WRAP_COLUMNS`). The three date columns (`DATE_DISPLAY_COLUMNS`) are real `datetime` values formatted `dd-mmm-yy` for display — critically, "Insuff Raised date" keeps its full timestamp in the underlying cell value; only the *display* is date-only. Don't truncate that column to a date string — the time component is intentionally preserved.

## Output files (written to the same directory as `app.py`)

- `Yesterday Done checks with other details as per client name - Bridge.csv` — raw daily export
- `Advance Tracker.csv`, `Antecedent Details.csv` — the two lookup exports
- `Accenture_Closed_with_Insuff_<date>.xlsx` — the final formatted deliverable

The three CSVs are the raw query results written by `save_csv` (kept for audit and the Streamlit download expander).
