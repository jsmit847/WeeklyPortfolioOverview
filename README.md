# Weekly Portfolio Meeting Deck

A Streamlit meeting workspace for Bridge and Term portfolios. Upload the weekly Excel workbook, review the current loan picture, capture edits and follow-ups, and download a verified workbook for the next meeting.

## Run locally

Use Python 3.11 or later:

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python -m streamlit run portfolio_meeting_app.py
```

Upload a workbook through the sidebar. For a local default, set `WPO_DEFAULT_WORKBOOK` to its path, or place `portfolio_overview.xlsx` in the working directory. Local inputs, credentials, caches and recovery files are excluded from Git.

## Meeting workflow

1. Check **Overview** for workbook readiness, missing owners, attention signals, balance composition and the maturity wall. Set the meeting date separately from the workbook data date.
2. Use **Presentation** to move through the agenda. Type, owner, search and focus filters carry between views. Choose workbook order, owner groups, attention first, maturity or balance order. The selected deal remains stable when filters change.
3. Read **As it stands today**, the source consistency checks and Land Gorilla panel. Expand workbook fields, comparison history and supplemental sheets when needed. Workbook and live values keep their source labels.
4. Edit status, owner or commentary from the loan editor, or commit individual cells in **Review**. Empty fields stay empty. Edits apply to the loaded workbook; loading a different file prompts a keep/discard decision.
5. Capture actions, decisions and notes, mark meeting outcomes, and review the open register in **Follow-ups**.
6. Review **Wrap-up**, download minutes or an Outlook draft, and use **Exports → Prepare and verify export**. Download the checked workbook and carry it forward into the next meeting.

Arrow keys navigate the presentation; **E** opens the editor. Keyboard shortcuts pause while typing. Compact presentation, pinned fields and agenda settings are available in the sidebar.

## Workbook contract

The reader recognizes Bridge and Term tabs, including normalized tab names. It searches the first 60 rows for a `Deal Number` header and supports the configured alternative loan-number header. Status, owner and AM Commentary are resolved consistently for both reading and writing. Duplicate deal numbers receive separate row keys. Hidden rows can be included explicitly; totals and Excel errors are excluded from the agenda.

Dated UPB and NPL columns are selected on or before the chosen data date. Each sheet reports which column it used. Missing or invalid dates, missing tabs, duplicate headers, formula cells without cached results and unavailable loan sizes are surfaced in workbook notes. Term commitment falls back to original loan amount when no usable commitment is available.

`Loan Modifications` is optional. Cancelled and inactive records remain inspectable but do not displace an active headline modification.

The `_Status History` schema remains:

```text
meeting_date, sheet, deal_number, deal_key, status, owner, commentary, workbook_name, saved_at
```

Application-owned meeting sheets are hidden in the exported workbook. Preparing again updates the current meeting rather than adding duplicate snapshots. Edits use OOXML cell replacements so untouched formulas, their cached values and workbook formatting survive export. Entered text beginning with `=` is stored as literal text. Generated CSV downloads neutralize spreadsheet formulas.

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

The app matches a Land Gorilla file number using the prefix plus the selected workbook loan-number column. The sidebar reports matched and unmatched loans, connection status and refresh controls. Panels show construction progress against schedule, balances, funding remaining, draw activity, due dates and risk labels. Disagreements with workbook amounts and dates are shown explicitly.

Successful responses are cached to limit API traffic. Login failures back off; errors shown to viewers contain exception types rather than credential-bearing response text. Without configuration or when the service is unavailable, workbook views remain usable. The integration tests use a fake service; no live credentials are required for tests.

Transient HTTP 500/502/503/504 and rate-limit responses receive up to three read attempts with bounded backoff. Requests use at most four workers. Persistent read errors pause automatic retries for one minute; Refresh retries immediately. A failed refresh retains the last successful values with an explicit last-known label and original fetch time. Partial responses identify the unavailable source fields.

## Bring more Salesforce data into the deck

Add visible worksheets containing a `Deal Number`, `Opportunity Deal Loan Number` or `Loan Number` header. Matching rows appear on the relevant loan, without a Salesforce connection. A single-row extract can carry stage, next step, latest activity and links; a multi-row sheet can carry draw or servicing history. HTTP(S) link fields provide direct source links. Keep extract dates in the workbook so viewers can judge freshness.

Box document discovery, when enabled, reads local file metadata from an explicitly configured folder. It does not read document contents. Box Drive must be available on the computer running Streamlit; a local desktop folder is not available to Community Cloud.

## Recovery and privacy

The recovery controls provide portable meeting backups and undo/redo for saved edits. Local disk autosave is opt-in and should point to a private folder; it is not a shared editing service. Recovery is checked against the workbook fingerprint before applying edits. Download a recovery file and the verified workbook before ending a meeting or restarting the hosting process.

The app does not send email. The Outlook download creates a draft `.eml` for review in the mail client. Meeting notes and exports can contain portfolio information; store them in the same controlled location as the source workbook.

## Checks and deployment

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest -q
python -m pytest -q -m slow
```

The suite uses synthetic workbooks and mocked Land Gorilla requests. CI checks Windows and Linux on Python 3.11–3.13, includes pandas 2 compatibility and runs weekly against the allowed dependency ranges. An optional browser smoke check uses Playwright with Edge: `python tests/browser_smoke.py`.

Streamlit Community Cloud can deploy `portfolio_meeting_app.py` from `main`. Configure credentials through the hosting service's secrets UI. The repository includes a 50 MB upload limit, viewer toolbar, restricted error details and disabled usage telemetry.
