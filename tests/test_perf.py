"""Performance budgets (slow; run with `pytest -m slow`).

- A workbook with 60,000 formatted-but-empty rows parses in under 4 s.
- At 120 Bridge + 30 Term deals x 52 weekly meetings, the vectorized streaks compute in
  under 0.5 s cold and agree with the old per-deal loop, and the median Next click under
  AppTest takes at most 1.0 s (about 3 s before the parse/derive split).
"""
from __future__ import annotations

import datetime as dt
import statistics
import time

import pytest

import sample_workbook as sw
from app_helpers import button_by_label, progress_text
from test_core import _reference_status_streaks

pytestmark = pytest.mark.slow

AS_OF = dt.date(2026, 10, 15)


def test_parse_skips_60000_styled_empty_rows(app_module):
    data = sw.with_styled_empty_rows(sw.build_sample_workbook_bytes(as_of=AS_OF), "Bridge", 60_000)
    start = time.perf_counter()
    parsed = app_module.parse_workbook_bytes(data)
    elapsed = time.perf_counter() - start
    assert parsed["error"] is None
    assert parsed["sheets"]["Bridge"]["blank_stop_row"] is not None
    deck, _ = app_module.build_deck(parsed, AS_OF, False)
    assert len(deck) == 10 and deck["upb"].sum() == 66_769_850
    assert elapsed < 4.0, f"parse took {elapsed:.2f}s"


def test_streaks_at_scale_are_fast_and_match_the_reference(app_module, scaled_bytes):
    parsed = app_module.parse_workbook_bytes(scaled_bytes)
    deck, _ = app_module.build_deck(parsed, AS_OF, False)
    assert len(deck) == 150
    history = app_module.embedded_status_history(parsed)
    assert len(history) == 150 * 52
    pairs = tuple(zip(deck["deal_key"], deck["status"]))
    start = time.perf_counter()
    table = app_module.status_streak_table(history, pairs, AS_OF.isoformat(), 3)
    elapsed = time.perf_counter() - start
    assert elapsed < 0.5, f"streaks took {elapsed:.2f}s"
    expected = _reference_status_streaks(history, list(pairs), AS_OF)
    for key, (count, since, prompt) in expected.items():
        row = table.loc[key]
        assert (int(row["status_same_meeting_count"]), row["status_same_since"], bool(row["status_needs_update_prompt"])) == (
            count, since, prompt), key


def test_next_click_median_under_one_second(make_at, scaled_bytes):
    at = make_at(scaled_bytes)
    at.run()
    assert not at.exception
    timings = []
    for _ in range(9):
        start = time.perf_counter()
        button_by_label(at, "Next").click().run()
        timings.append(time.perf_counter() - start)
    assert not at.exception
    assert progress_text(at) == "Agenda item 10 of 150"
    median = statistics.median(timings)
    assert median <= 1.0, f"median Next click {median:.2f}s ({[round(t, 2) for t in timings]})"
