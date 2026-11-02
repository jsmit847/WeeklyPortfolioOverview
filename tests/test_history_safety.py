"""Status-history safety: a CSV upload is validated and merged, never a replacement (U04),
and exports leave the generator's _Status History sheet alone unless history changed (U28)."""
import datetime as dt
import io
from types import SimpleNamespace

import pandas as pd
import pytest
from openpyxl import load_workbook

import portfolio_exports as exports
import sample_workbook as sw
from conftest import AS_OF


class State(dict):
    def __getattr__(self, name):
        try:
            return self[name]
        except KeyError as exc:
            raise AttributeError(name) from exc

    def __setattr__(self, name, value):
        self[name] = value


class Upload:
    def __init__(self, name, data):
        self.name, self._data = name, data

    def getvalue(self):
        return self._data


def context(app, data):
    parsed = app.parse_workbook_bytes(data)
    deck, _info = app.build_deck(parsed, AS_OF, False)
    return SimpleNamespace(wb_sha=app.bytes_sha256(data), wb={"bytes": data}, workbook_name="synthetic.xlsx",
                           as_of=AS_OF, raw_deck=deck, deck=deck, parsed=parsed)


def csv_bytes(frame):
    return frame.to_csv(index=False).encode("utf-8")


@pytest.fixture
def state(app_module, monkeypatch):
    session = State()
    monkeypatch.setattr(app_module.st, "session_state", session)
    app_module.init_state()
    return session


def load_workbook_history(app, state, data, name="week.xlsx"):
    app.load_embedded_status_history_if_available(app.parse_workbook_bytes(data), app.bytes_sha256(data), name)
    return state.status_history


def test_agenda_csv_upload_is_rejected_and_history_kept(app_module, state, sample_bytes):
    app = app_module
    history = load_workbook_history(app, state, sample_bytes)
    assert len(history) == 21
    agenda = app.build_agenda_dataframe(context(app, sample_bytes).deck)
    with pytest.raises(ValueError, match="Not a status-history file"):
        app.history_upload_frame(csv_bytes(agenda))
    state.uploaded_status_history = Upload("weekly_portfolio_agenda.csv", csv_bytes(agenda))
    app.process_status_history_upload()
    assert len(state.status_history) == 21
    assert "Not a status-history file" in state.status_history_load_error


def test_older_backup_merges_instead_of_replacing(app_module, state, sample_bytes):
    app = app_module
    history = load_workbook_history(app, state, sample_bytes)
    oldest = history[history.meeting_date == history.meeting_date.min()]
    state.uploaded_status_history = Upload("old_backup.csv", csv_bytes(oldest))
    app.process_status_history_upload()
    assert len(state.status_history) == 21  # all three meetings survive
    assert set(state.status_history.meeting_date) == set(history.meeting_date)
    assert state.status_history_load_error == ""


def test_later_saved_at_wins_and_ties_keep_existing(app_module):
    app = app_module
    base = {"meeting_date": "2026-10-08", "sheet": "Bridge", "deal_number": "1", "deal_key": "Bridge::1", "owner": "", "commentary": "", "workbook_name": ""}
    existing = pd.DataFrame([{**base, "status": "Watchlist", "saved_at": "2026-10-08 10:00:00"}])
    newer = pd.DataFrame([{**base, "status": "Resolved", "saved_at": "2026-10-09 09:00:00"}])
    tie = pd.DataFrame([{**base, "status": "Other", "saved_at": "2026-10-08 10:00:00"}])
    merged, used = app.merge_status_history_frames(existing, newer)
    assert list(merged.status) == ["Resolved"] and used == 1
    merged, used = app.merge_status_history_frames(existing, tie)
    assert list(merged.status) == ["Watchlist"] and used == 0
    new_meeting = pd.DataFrame([{**base, "meeting_date": "2026-10-15", "status": "Watchlist", "saved_at": ""}])
    merged, used = app.merge_status_history_frames(existing, new_meeting)
    assert len(merged) == 2 and used == 1


def test_csv_with_sheet_and_deal_number_but_no_deal_key(app_module):
    data = b"meeting_date,sheet,deal_number,status\n2026-10-08,Bridge,10001,Surveillance\n2026-10-08,Term,0012345,Watchlist\n"
    frame = app_module.history_upload_frame(data)
    assert list(frame.deal_key) == ["Bridge::10001", "Term::0012345"]


def test_bom_csv_from_excel_is_accepted(app_module):
    data = "﻿meeting_date,deal_key,status\n10/08/2026,Bridge::10001,Surveillance\n".encode("utf-8")
    frame = app_module.history_upload_frame(data)
    assert list(frame.meeting_date) == ["2026-10-08"]


def test_removing_the_csv_reverts_to_the_workbook_history(app_module, state, sample_bytes):
    app = app_module
    load_workbook_history(app, state, sample_bytes)
    extra = pd.DataFrame([{"meeting_date": "2026-10-15", "sheet": "Bridge", "deal_number": "10001", "deal_key": "Bridge::10001", "status": "Resolved"}])
    state.uploaded_status_history = Upload("extra.csv", csv_bytes(extra))
    app.process_status_history_upload()
    assert len(state.status_history) == 22
    state.uploaded_status_history = None
    app.process_status_history_upload()
    assert len(state.status_history) == 21


def generator_style_workbook():
    """History sheet as an external generator might leave it: veryHidden, an extra column
    and a row the app can't read."""
    wb = sw.build_sample_workbook(as_of=AS_OF)
    ws = wb["_Status History"]
    ws.sheet_state = "veryHidden"
    ws.cell(1, 10).value = "upb"
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 10).value = 1000 * r
    ws.append(["not a date", "Bridge", "10001", "Bridge::10001", "", "", "", "", "", 7])
    return sw.workbook_to_bytes(wb)


def history_sheet(data):
    wb = load_workbook(io.BytesIO(data), read_only=True)
    try:
        ws = wb["_Status History"]
        return ws.sheet_state, [tuple(row) for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def test_export_without_history_change_leaves_the_sheet_untouched(app_module):
    app = app_module
    data = generator_style_workbook()
    ctx = context(app, data)
    state = {"meeting_overrides": {}, "status_history": app.embedded_status_history(ctx.parsed)}
    bundle = exports.prepare_export(ctx, app, state)
    before_state, before_rows = history_sheet(data)
    after_state, after_rows = history_sheet(bundle["workbook"])
    assert after_state == before_state == "veryHidden"
    assert after_rows == before_rows  # extra column and unreadable row preserved


def test_export_after_a_csv_merge_rewrites_history_and_keeps_it_hidden(app_module, state):
    app = app_module
    data = generator_style_workbook()
    load_workbook_history(app, state, data)
    extra = pd.DataFrame([{"meeting_date": "2026-10-15", "sheet": "Bridge", "deal_number": "10001", "deal_key": "Bridge::10001", "status": "Resolved"}])
    state.uploaded_status_history = Upload("extra.csv", csv_bytes(extra))
    app.process_status_history_upload()
    ctx = context(app, data)
    bundle = exports.prepare_export(ctx, app, {"meeting_overrides": {}, "status_history": state.status_history})
    sheet_state, rows = history_sheet(bundle["workbook"])
    assert sheet_state == "veryHidden"
    assert list(rows[0]) == app.STATUS_HISTORY_COLUMNS
    assert len(rows) - 1 == len(state.status_history) == 22


@pytest.mark.parametrize("name, expected", [
    ("Portfolio Overview 10.15.26.xlsx", "Portfolio Overview 10.15.26_meeting_2026-11-02.xlsx"),
    ("Portfolio Overview 10.15.26_meeting_2026-10-26.xlsx", "Portfolio Overview 10.15.26_meeting_2026-11-02.xlsx"),
    ("X_meeting_2026-10-19_meeting_2026-10-26.xlsx", "X_meeting_2026-11-02.xlsx"),
])
def test_export_file_name_does_not_stack(name, expected):
    assert exports.export_filename(name, dt.date(2026, 11, 2)) == expected
