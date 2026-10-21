# Weekly Portfolio Meeting Deck

A Streamlit deck for Bridge and Term portfolio meetings. Upload the weekly Excel workbook, cycle through loans, edit status or commentary when needed, and download the checked workbook.

## Run locally

Use Python 3.11 or later:

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python -m streamlit run portfolio_meeting_app.py
```

Upload a workbook through the sidebar. For a local default, set `WPO_DEFAULT_WORKBOOK` to its path, or place `portfolio_overview.xlsx` in the working directory. Local inputs, credentials, caches and recovery files are excluded from Git.

## Meeting workflow

1. Use **Overview** for balances, maturities and the loan modification list.
2. Use **Presentation** to move through loans with Previous/Next or the jump picker. Arrow keys navigate; **E** opens the editor. Filters and the selected loan persist between views.
3. Read the workbook KPIs, deal snapshot and AM commentary. Optional construction data, loan modifications and additional workbook fields are collapsed. Open **Construction · Land Gorilla** and click **Load construction data** only when needed.
4. Save status, owner or commentary edits in the loan editor or **Review**. Notes and attendance are optional in the meeting itself; the app does not require a capture, follow-up or wrap-up workflow.
5. Use **Exports → Prepare and verify export**, then download the checked workbook. Only saved edits are applied. Existing status history is preserved; exporting does not add a snapshot for every loan.

Historical update reminders are off by default. They can be enabled under Presentation settings, but repeated workbook text does not establish whether a loan was discussed. The app does not create a meeting journal, timer, action register or meeting recovery link. Download your edited workbook before ending the session.

## Workbook contract

The reader recognizes Bridge and Term tabs, including normalized tab names. It searches the first 60 rows for a `Deal Number` header and supports the configured alternative loan-number header. Status, owner and AM Commentary are resolved consistently for both reading and writing. Duplicate deal numbers receive separate row keys. Hidden rows can be included explicitly; totals and Excel errors are excluded from the agenda.

Dated UPB and NPL columns are selected on or before the chosen data date. Each sheet reports which column it used. Missing or invalid dates, missing tabs, duplicate headers, formula cells without cached results and unavailable loan sizes are surfaced in workbook notes. Term commitment falls back to original loan amount when no usable commitment is available.

`Loan Modifications` is optional. Cancelled and inactive records remain inspectable but do not displace an active headline modification.

The `_Status History` schema remains:

```text
meeting_date, sheet, deal_number, deal_key, status, owner, commentary, workbook_name, saved_at
```

Existing hidden sheets are preserved. No follow-up, progress or minutes sheets are generated. Edits use OOXML cell replacements so untouched formulas, their cached values and workbook formatting survive export. Entered text beginning with `=` is stored as literal text. Generated CSV downloads neutralize spreadsheet formulas.

## Land Gorilla

Bridge construction data is optional. Configure the deployed app's Streamlit secrets, or `.streamlit/secrets.toml` locally:

```toml
[landgorilla]
user = "YOUR_USER"
password = "YOUR_PASSWORD"
verify = true
# template_id = "437"
# file_prefix = "rb0"
```

The app matches a Land Gorilla file number using the prefix plus the selected workbook loan-number column. Requests use the same token, filtered loan lookup, pipeline template and loan-summary endpoints shown in the supplied notebook. Explicit file-number matches take priority; ambiguous responses are rejected.

**No API requests run during slide navigation, Review, Overview or Exports.** Load/Refresh requests only the selected Bridge loan. The collapsed panel shows construction and schedule percentages, balances, holdback remaining, approved draw date and LG due date. LG labels and instructions are neutral source metadata. Unavailable data is described inside the panel, without a page-level alert. Workbook values remain available independently.

Saved snapshots retain their original retrieval time. Expired snapshots and failed refreshes are identified rather than represented as current. Requests retry transient server errors and rate limits with bounded backoff. Errors contain exception types/status codes, never raw response bodies or credential-bearing exception text.

See [the notebook findings and safe local diagnostic](docs/data-sources.md) for the available LG and FCI fields, observed endpoint permissions and a command to diagnose one affected loan. FCI is not connected live in this app; its current servicing figures come from the workbook.

## Bring more Salesforce data into the deck

Add visible worksheets containing a `Deal Number`, `Opportunity Deal Loan Number` or `Loan Number` header. Matching rows appear on the relevant loan, without a Salesforce connection. A single-row extract can carry stage, next step, latest activity and links; a multi-row sheet can carry draw or servicing history. HTTP(S) link fields provide direct source links. Keep extract dates in the workbook so viewers can judge freshness.

Box document discovery, when enabled, reads local file metadata from an explicitly configured folder. It does not read document contents. Box Drive must be available on the computer running Streamlit; a local desktop folder is not available to Community Cloud.

## Session edits

Saved edits stay in the Streamlit session. Export the workbook before restarting the app or ending the session. Loading a different workbook protects edits with an explicit keep/discard decision. Credentials and local portfolio inputs are excluded from Git.

## Checks and deployment

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest -q
python -m pytest -q -m slow
```

The suite uses synthetic workbooks and mocked Land Gorilla requests. CI checks Windows and Linux on Python 3.11–3.13, includes pandas 2 compatibility and runs weekly against the allowed dependency ranges. An optional browser smoke check uses Playwright with Edge: `python tests/browser_smoke.py`.

Streamlit Community Cloud can deploy `portfolio_meeting_app.py` from `main`. Configure credentials through the hosting service's secrets UI. The repository includes a 50 MB upload limit, viewer toolbar, restricted error details and disabled usage telemetry.
