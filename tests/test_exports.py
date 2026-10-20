import datetime as dt
from types import SimpleNamespace

import pandas as pd

import portfolio_exports as exports


def context(app, data):
    parsed = app.parse_workbook_bytes(data)
    deck, info = app.build_deck(parsed, dt.date(2026, 10, 15), False)
    return SimpleNamespace(wb_sha=app.bytes_sha256(data), wb={"bytes": data}, workbook_name="synthetic.xlsx",
                           as_of=dt.date(2026, 10, 15), raw_deck=deck, deck=deck, parsed=parsed)


def test_prepare_export_round_trip(app_module, sample_bytes):
    ctx = context(app_module, sample_bytes)
    state = {"meeting_overrides": {}, "status_history": app_module.embedded_status_history(ctx.parsed)}
    bundle = exports.prepare_export(ctx, app_module, state)
    assert bundle["verification"]["rows"] >= 10
    result = app_module.parse_workbook_bytes(bundle["workbook"])
    history = app_module.embedded_status_history(result)
    assert "2026-10-15" in history.meeting_date.tolist()
    assert bundle["fingerprint"] == exports.fingerprint(ctx, state)


def test_fingerprint_changes_when_meeting_work_changes(app_module, sample_bytes):
    ctx = context(app_module, sample_bytes)
    state = {"meeting_items": []}
    before = exports.fingerprint(ctx, state)
    state["meeting_items"] = [{"text": "Synthetic action", "owner": "Alice"}]
    assert before != exports.fingerprint(ctx, state)


def test_repeat_export_does_not_duplicate_history(app_module, sample_bytes):
    ctx = context(app_module, sample_bytes)
    state = {"meeting_overrides": {}, "status_history": app_module.embedded_status_history(ctx.parsed)}
    first = exports.prepare_export(ctx, app_module, state)
    state["status_history"] = app_module.read_status_history_csv(__import__('io').BytesIO(first["history"]))
    second = exports.prepare_export(ctx, app_module, state)
    first_frame = app_module.read_status_history_csv(__import__('io').BytesIO(first["history"]))
    second_frame = app_module.read_status_history_csv(__import__('io').BytesIO(second["history"]))
    assert len(first_frame) == len(second_frame)
    assert not second_frame.duplicated(["meeting_date", "deal_key"]).any()
