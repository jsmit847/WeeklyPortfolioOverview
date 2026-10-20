"""Unit tests for the pure helpers: text, money, dates, dated headers, note splitting,
maturity buckets, timing signals, layouts and the vectorized status streaks."""
from __future__ import annotations

import datetime as dt
import math
import random
import re

import numpy as np
import pandas as pd
import pytest

AS_OF = dt.date(2026, 10, 15)


# -----------------------------------------------------------------------------
# Text helpers
# -----------------------------------------------------------------------------
def test_norm_label_keeps_text_as_typed(app_module):
    assert app_module.norm_label("Rate step-up to 9.0") == "Rate step-up to 9.0"
    assert app_module.norm_label("  Upcoming\n Maturity  ") == "Upcoming Maturity"
    assert app_module.norm_label("Phase\xa02") == "Phase 2"
    assert app_module.norm_label(9.0) == "9"
    assert app_module.norm_label(None) == ""
    assert app_module.norm_label(float("nan")) == ""
    assert app_module.norm_label(pd.NA) == ""
    assert app_module.norm_label(dt.datetime(2026, 10, 15)) == "10/15/2026"


def test_norm_id_rules(app_module):
    assert app_module.norm_id(10001.0) == "10001"
    assert app_module.norm_id("10001.0") == "10001"
    assert app_module.norm_id("00123") == "00123"
    assert app_module.norm_id(" 10001 ") == "10001"
    assert app_module.norm_id(10001) == "10001"
    assert app_module.norm_id(np.int64(10001)) == "10001"
    assert app_module.norm_id("10001.5") == "10001.5"
    assert app_module.norm_id("BR-10009") == "BR-10009"
    assert app_module.norm_id(None) == ""
    assert app_module.norm_id(float("nan")) == ""


def test_clean_free_text_keeps_line_breaks(app_module):
    clean = app_module.clean_free_text
    assert clean("a\r\nb\rc") == "a\nb\nc"
    assert clean("a\x0bb") == "a\nb"
    assert clean("a\x0cb") == "a\nb"
    assert clean("a\x07b") == "a b"
    assert clean("x  \ny\t") == "x\ny"
    assert clean("a\n\n\n\nb") == "a\n\nb"
    assert clean("Rate step-up to 9.0") == "Rate step-up to 9.0"
    assert clean("Line one_x000D_\nLine two") == "Line one\nLine two"
    assert clean("  leading kept inside\n  indented") == "leading kept inside\n  indented"
    assert clean(None) == "" and clean(float("nan")) == ""


def test_text_equal_is_whitespace_insensitive(app_module):
    assert app_module.text_equal("a\nb", "a  b") is True
    assert app_module.text_equal("Line one\nLine two", "Line one Line two") is True
    assert app_module.text_equal("Rate step-up to 9.0", "Rate step-up to 9") is True  # legacy rows
    assert app_module.text_equal("a", "b") is False


def test_keys(app_module):
    assert app_module.deal_match_key("00010002") == app_module.deal_match_key(10002) == "10002"
    assert app_module.deal_match_key("10002.0") == "10002"
    assert app_module.deal_match_key("Br-1") == app_module.deal_match_key("BR-1")
    assert app_module.norm_deal_key("Bridge::10001.0") == "Bridge::10001"
    assert app_module.norm_deal_key("Bridge::10002#r14") == "Bridge::10002#r14"
    assert app_module.get_override_key("Bridge", 10001.0) == "Bridge::10001"


# -----------------------------------------------------------------------------
# Money
# -----------------------------------------------------------------------------
@pytest.mark.parametrize(
    "value, expected",
    [
        (-1_500_000, "-$1.5MM"),
        (999_999.6, "$1.0MM"),
        (999_999.4, "$999,999"),
        (-0.0, "$0"),
        (-4.656612873077393e-10, "$0"),
        (float("inf"), "-"),
        (float("-inf"), "-"),
        (float("nan"), "-"),
        (None, "-"),
        (pd.NA, "-"),
        ("abc", "-"),
        (-250_000, "-$250,000"),
        (22_000_000, "$22.0MM"),
        (12_500_000, "$12.5MM"),
        (np.float64(10_250_000), "$10.2MM"),
    ],
)
def test_fmt_money(app_module, value, expected):
    assert app_module.fmt_money(value) == expected


def test_fmt_money_exact(app_module):
    assert app_module.fmt_money_exact(10_250_000) == "$10,250,000"
    assert app_module.fmt_money_exact(-250_000) == "-$250,000"
    assert app_module.fmt_money_exact(-0.4) == "$0"
    assert app_module.fmt_money_exact(13_540_000) != app_module.fmt_money_exact(13_500_000)
    assert app_module.fmt_money_exact(None) == "-"


def test_coerce_number_reads_currency_text(app_module):
    assert app_module.coerce_number("$1,250,000") == 1_250_000
    assert app_module.coerce_number("(1,000)") == -1000
    assert math.isnan(app_module.coerce_number("TBD"))
    assert math.isnan(app_module.coerce_number(True))


# -----------------------------------------------------------------------------
# Dates
# -----------------------------------------------------------------------------
def _ts(text):
    return pd.Timestamp(text)


@pytest.mark.parametrize(
    "value, expected",
    [
        (dt.datetime(2026, 10, 15, 9, 30), _ts("2026-10-15 09:30")),
        (dt.date(2026, 10, 15), _ts("2026-10-15")),
        (pd.Timestamp("2026-10-15"), _ts("2026-10-15")),
        (45000, _ts("2023-03-15")),
        (45000.5, _ts("2023-03-15 12:00")),
        ("45000", _ts("2023-03-15")),
        (20260331, _ts("2026-03-31")),
        ("2026-10-30T00:00:00.000+0000", _ts("2026-10-30")),
        ("2026-09-05T17:22:10.000+0000", _ts("2026-09-05 17:22:10")),
        ("2026-09-28 14:22:10 UTC", _ts("2026-09-28 14:22:10")),
        ("10/15/2026", _ts("2026-10-15")),
        ("2026-10-15", _ts("2026-10-15")),
        (5, pd.NaT),
        (2026, pd.NaT),
        (True, pd.NaT),
        (None, pd.NaT),
        ("", pd.NaT),
        ("TBD", pd.NaT),
        ("10/15", pd.NaT),
        ("12/31/9999", pd.NaT),
        (dt.datetime(3026, 6, 1), pd.NaT),
        (dt.datetime(9999, 12, 31), pd.NaT),
        (dt.datetime(1899, 1, 1), pd.NaT),
    ],
)
def test_coerce_datetime_value(app_module, value, expected):
    result = app_module.coerce_datetime_value(value)
    if pd.isna(expected):
        assert pd.isna(result)
    else:
        assert pd.Timestamp(result) == expected


def test_coerce_datetime_series_never_raises_and_is_naive(app_module):
    mixed = pd.Series(
        [
            dt.datetime(2026, 10, 15),
            "2026-10-14T21:05:33-0700",
            "2026-10-14T21:05:33+0000",
            45000,
            "12/31/9999",
            dt.datetime(3026, 6, 1),
            None,
            "TBD",
            True,
            "2026-09-28 14:22:10 UTC",
        ],
        dtype=object,
    )
    parsed, rejected = app_module._coerce_datetime_core(mixed)
    assert str(parsed.dtype) == "datetime64[ns]"
    assert getattr(parsed.dt, "tz", None) is None
    assert parsed.iloc[1] == _ts("2026-10-15 04:05:33")  # converted to UTC wall time
    assert parsed.iloc[3] == _ts("2023-03-15")
    assert rejected.tolist() == [False, False, False, False, True, True, False, True, True, False]
    # all-UTC column (the Salesforce API format) and tz-aware datetime64 input
    utc_only = pd.Series(["2026-09-05T17:22:10.000+0000"] * 3, dtype=object)
    assert app_module.coerce_datetime_series(utc_only).notna().all()
    aware = pd.Series(pd.to_datetime(["2026-10-15 12:00"]).tz_localize("UTC"))
    assert str(app_module.coerce_datetime_series(aware).dtype) == "datetime64[ns]"
    # a serial column and a 3026 column concatenate without OutOfBoundsDatetime
    a = app_module.coerce_datetime_series(pd.Series([45000, dt.datetime(2026, 1, 1)], dtype=object))
    b = app_module.coerce_datetime_series(pd.Series([dt.datetime(3026, 6, 1)], dtype=object))
    combined = pd.concat([a, b], ignore_index=True)
    assert str(combined.dtype) == "datetime64[ns]" and combined.isna().tolist() == [False, False, True]


def test_fmt_date_never_shows_1970(app_module):
    assert app_module.fmt_date(5) == "-"
    assert app_module.fmt_date(45000) == "03/15/2023"
    assert app_module.fmt_date(2026) == "-"
    for value in (0, 1, 5, 2026, 20260331, "0", 1e12):
        assert app_module.fmt_date(value) != "01/01/1970"


# -----------------------------------------------------------------------------
# Dated headers ('10/15 UPB') and the as-of choice
# -----------------------------------------------------------------------------
def test_parse_header_date_closest_year(app_module):
    parse = app_module.parse_header_date
    assert parse(12, 31, None, dt.date(2027, 1, 7)) == dt.date(2026, 12, 31)
    assert parse(1, 7, None, dt.date(2026, 12, 31)) == dt.date(2027, 1, 7)
    assert parse(10, 15, None, AS_OF) == dt.date(2026, 10, 15)
    assert parse(10, 15, "26", AS_OF) == dt.date(2026, 10, 15)
    assert parse(2, 29, None, dt.date(2028, 2, 28)) == dt.date(2028, 2, 29)
    assert parse(2, 30, None, AS_OF) is None


def test_latest_dated_by_suffix_uses_as_of(app_module):
    headers = ["10/8 UPB", "10/15 UPB", "10/8 NPL", "10/15 NPL"]
    pick = app_module.latest_dated_by_suffix
    assert pick(headers, as_of_date=AS_OF, reference_date=AS_OF) == {"UPB": "10/15 UPB", "NPL": "10/15 NPL"}
    assert pick(headers, as_of_date=dt.date(2026, 10, 14), reference_date=AS_OF) == {"UPB": "10/8 UPB", "NPL": "10/8 NPL"}
    # year boundary: '12/31' next to '1/7/27'
    boundary = ["12/31 UPB", "1/7/27 UPB"]
    assert pick(boundary, as_of_date=dt.date(2027, 1, 7), reference_date=dt.date(2027, 1, 7)) == {"UPB": "1/7/27 UPB"}
    assert pick(boundary, as_of_date=dt.date(2027, 1, 6), reference_date=dt.date(2027, 1, 7)) == {"UPB": "12/31 UPB"}
    # a single forward-dated column (quarter-end NPL label) is still used
    assert pick(["9/30 NPL"], as_of_date=dt.date(2026, 7, 3), reference_date=dt.date(2026, 7, 3)) == {"NPL": "9/30 NPL"}


# -----------------------------------------------------------------------------
# Loan-mod note splitting (U03)
# -----------------------------------------------------------------------------
@pytest.mark.parametrize(
    "text",
    [
        "Rate cut -0.50% and reserve top-up.",
        "Floor -0.25% per IC",
        "-$12,345.67 credit",
        "Unit (V) kitchen remodel",
        "1.5% exit fee waived",
        "2.0x DSCR covenant reset",
        "Extended maturity to 2026. Borrower paid fee",
        "Cedar Point Ventures (LLC) signed",
        "Borrower (MD) and (DC) entities; see (c)",
        "Units 1 - 4 renovated; rate 9% - 10% band",
        "$3.4MM - $3.9MM range",
        "-0.50% rate reduction approved",
        "Paid 1. Fee",
        "DSCR covenant reset to 1.0",
    ],
)
def test_split_keeps_non_list_text_unchanged(app_module, text):
    assert app_module.split_loan_mod_note_lines("Comments", text) == [text]


@pytest.mark.parametrize(
    "text, expected",
    [
        (
            "1. Rate step-up approved 2. Sponsor guaranty reaffirmed 3. Cash management triggered",
            ["Rate step-up approved", "Sponsor guaranty reaffirmed", "Cash management triggered"],
        ),
        (
            "Initial extension (i) paid down $100k (ii) added interest reserve",
            ["Initial extension", "paid down $100k", "added interest reserve"],
        ),
        ("Line one\nLine two", ["Line one", "Line two"]),
        ("1) Extended maturity\n2) Rate change\n3) Fee waived", ["Extended maturity", "Rate change", "Fee waived"]),
        ("• Item A • Item B", ["Item A", "Item B"]),
        ("- First\n- Second", ["First", "Second"]),
        ("Extended twice: 1. to 2025. 2. to 2026.", ["Extended twice:", "to 2025.", "to 2026."]),
        (
            "Comments: Extended maturity 6 months to 2026. Borrower paid 0.5% extension fee - Updated budget approved",
            ["Extended maturity 6 months to 2026. Borrower paid 0.5% extension fee", "Updated budget approved"],
        ),
        ("Forbearance through 4/30.\nRate reset to 10.0", ["Forbearance through 4/30.", "Rate reset to 10.0"]),
    ],
)
def test_split_real_lists(app_module, text, expected):
    assert app_module.split_loan_mod_note_lines("Comments", text) == expected


_NUMBER_TOKEN = re.compile(r"-?\$?\d[\d,]*(?:\.\d+)?%?x?|\([A-Z][A-Za-z]*\)")


@pytest.mark.parametrize("seed", range(40))
def test_split_never_drops_numbers_or_tags(app_module, seed):
    rng = random.Random(seed)
    words = ["Rate", "cut", "-0.50%", "fee", "$1,250,000", "(LLC)", "to", "2026.", "Borrower", "1.5%",
             "-$12,345.67", "(V)", "paid", "-", "x2.0", "DSCR", "1.35x"]
    text = " ".join(rng.choice(words) for _ in range(rng.randrange(3, 14)))
    joined = " ".join(app_module.split_loan_mod_note_lines("Comments", text))
    for token in _NUMBER_TOKEN.findall(text):
        assert token in joined, (text, token, joined)


def test_note_html_keeps_trailing_zero(app_module):
    html_block = app_module.loan_mod_note_section_html("Comments", "Forbearance through 4/30.\nRate reset to 10.0")
    assert "<li>Rate reset to 10.0</li>" in html_block


# -----------------------------------------------------------------------------
# Maturity buckets and timing pills (U43/U44)
# -----------------------------------------------------------------------------
def test_maturity_buckets(app_module):
    assert app_module.maturity_buckets(pd.Series([-120, -10, -1, 0, 12, 30, 31, 45, np.nan])) == (3, 3)
    assert app_module.maturity_buckets(pd.Series([np.nan, np.nan])) == (0, 0)


@pytest.mark.parametrize(
    "days, expected",
    [(-1, "rt-signal-red"), (0, "rt-signal-amber"), (1, "rt-signal-amber"), (12, "rt-signal-amber"),
     (30, "rt-signal-amber"), (31, "rt-signal-green"), (365, "rt-signal-green"), (None, "rt-signal-muted"),
     (float("nan"), "rt-signal-muted"), ("abc", "rt-signal-muted")],
)
def test_maturity_signal(app_module, days, expected):
    assert app_module.timing_signal(days, warn_within=app_module.MATURITY_WATCH_DAYS)[1] == expected


@pytest.mark.parametrize("days, expected", [(-3, "rt-signal-red"), (0, "rt-signal-amber"), (1, "rt-signal-green"), (25, "rt-signal-green")])
def test_payment_signal_unchanged(app_module, days, expected):
    assert app_module.timing_signal(days)[1] == expected


def test_maturity_signal_is_monotonic(app_module):
    rank = {"rt-signal-red": 2, "rt-signal-amber": 1, "rt-signal-green": 0}
    previous = 3
    for days in range(-400, 2000):
        current = rank[app_module.timing_signal(days, warn_within=30)[1]]
        assert current <= previous
        previous = current


# -----------------------------------------------------------------------------
# Sheet names and layouts
# -----------------------------------------------------------------------------
def test_resolve_sheet_name(app_module):
    resolve = app_module.resolve_sheet_name
    assert resolve(["Bridge", "TERM "], "Term") == ("TERM ", "normalized")
    assert resolve(["Bridge", "Term", "Term "], "Term") == ("Term", "exact")
    assert resolve(["TERM", "term "], "Term") == (None, "ambiguous")
    assert resolve(["Bridge"], "Term") == (None, "missing")
    assert resolve(["Loan modifications"], "Loan Modifications") == ("Loan modifications", "normalized")


def test_layout_duplicates_and_owner_priority(app_module):
    headers = ["Deal Number", "Active RM", "Status", "Point Person", "AM Commentary", "Status",
               "Days Past Due", "Days Past Due ", None, "10/15 UPB"]
    layout = app_module.layout_from_header_values(headers, header_row=4)
    assert layout["roles"]["status"] == "Status"
    assert layout["role_cols"]["status"] == 3
    assert layout["labels"][5] == "Status (2)"
    assert layout["role_cols"]["owner"] == 4  # Point Person beats a left-most Active RM
    assert layout["dpd_cols"] == [7, 8]
    assert layout["max_col"] == 10
    assert [d["label"] for d in layout["dated"]] == ["10/15 UPB"]
    assert layout["role_duplicates"] == [
        {"role": "status", "header": "Status", "columns": ["C", "F"], "used": "C", "coalesced": False}
    ]


@pytest.mark.parametrize(
    "value, expected",
    [("Yes", True), (" yes ", True), ("Y", True), ("TRUE", True), ("1", True), ("X", True), (True, True), (1, True),
     (1.0, True), ("No", False), ("N", False), (False, False), (0, False), ("FALSE", False), ("", None), (None, None),
     ("Pending", None)],
)
def test_parse_loan_mod_match_flag(app_module, value, expected):
    assert app_module.parse_loan_mod_match_flag(value) is expected


# -----------------------------------------------------------------------------
# Status streaks: vectorized vs the old loop
# -----------------------------------------------------------------------------
def _legacy_norm_text(value) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    text = re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()
    return text[:-2] if text.endswith(".0") else text


def _robust_date(value):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(value))
    return pd.Timestamp(value)


def _reference_status_streaks(history: pd.DataFrame, current_pairs, meeting_date: dt.date, threshold: int = 3):
    """The pre-P1 algorithm (one loop per deal over the merged history), kept as an oracle.

    Merge: history rows, then this meeting's snapshot (non-blank statuses); the last row
    per (meeting date, deal key) wins. Per deal: count consecutive most-recent rows whose
    status equals the current one; 'since' is the oldest date in that run.
    """
    rows = []
    for record in history.to_dict("records"):
        rows.append((_robust_date(record["meeting_date"]).normalize(), _legacy_norm_text(record["deal_key"]),
                     _legacy_norm_text(record["status"])))
    for key, status in current_pairs:
        if _legacy_norm_text(status):
            rows.append((pd.Timestamp(meeting_date), _legacy_norm_text(key), _legacy_norm_text(status)))
    merged = {}
    for day, key, status in rows:
        if key and status:
            merged[(day, key)] = status
    out = {}
    for key, status in current_pairs:
        key_n, status_n = _legacy_norm_text(key), _legacy_norm_text(status)
        count, since = 0, ""
        if key_n and status_n:
            deal_rows = sorted(((day, s) for (day, k), s in merged.items() if k == key_n), reverse=True)
            for day, s in deal_rows:
                if s != status_n:
                    break
                count += 1
                since = day.strftime("%m/%d/%Y")
        out[key] = (count, since, count >= threshold)
    return out


def _random_history(rng: random.Random):
    meeting = AS_OF
    n_deals = rng.randrange(3, 12)
    n_meetings = rng.randrange(1, 9)
    keys = [f"{rng.choice(['Bridge', 'Term'])}::{10000 + i}" for i in range(n_deals)]
    statuses = ["Surveillance", "Foreclosure", "Modification", "Phase 2.0", "Upcoming  Maturity"]
    dates = [meeting - dt.timedelta(days=7 * k) for k in range(1, n_meetings + 1)]
    if rng.random() < 0.3:
        dates.append(meeting)  # an earlier export on the meeting day itself
    rows = []
    for key in keys:
        status = rng.choice(statuses)
        for day in sorted(dates):
            if rng.random() < 0.3:
                status = rng.choice(statuses)
            fmt = rng.choice(["iso", "us", "dt", "serial"])
            if fmt == "iso":
                cell = day.isoformat()
            elif fmt == "us":
                cell = day.strftime("%m/%d/%Y")
            elif fmt == "dt":
                cell = dt.datetime.combine(day, dt.time())
            else:
                cell = (day - dt.date(1899, 12, 30)).days
            rows.append({"meeting_date": cell, "sheet": key.split("::")[0], "deal_number": key.split("::")[1],
                         "deal_key": key, "status": status, "owner": "", "commentary": "", "workbook_name": "",
                         "saved_at": ""})
            if rng.random() < 0.1:  # a duplicate same-date row (the later one wins)
                rows.append(dict(rows[-1], status=rng.choice(statuses)))
    if rng.random() < 0.5:
        rng.shuffle(rows)
    current = [(key, rng.choice(statuses + [""])) for key in keys]
    return pd.DataFrame(rows), current


def _reference_message(count: int, since: str, status: str) -> str:
    """The pre-P1 prompt wording (build_status_prompt_message), status shown as typed."""
    return (
        f"This deal has carried the same meeting update for {count} meetings since {since or 'unknown'}. "
        f"Confirm whether '{status or 'blank'}' still reflects the latest story, or refresh the update/commentary."
    )


def _assert_streaks_match(app_module, history_raw, current, meeting=AS_OF):
    normalized = app_module.normalize_status_history(history_raw)
    table = app_module.status_streak_table(normalized, tuple(current), meeting.isoformat(), 3)
    expected = _reference_status_streaks(history_raw, current, meeting)
    statuses = dict(current)
    for key, (count, since, prompt) in expected.items():
        got = (int(table.loc[key, "status_same_meeting_count"]) if key in table.index else 0,
               table.loc[key, "status_same_since"] if key in table.index else "",
               bool(table.loc[key, "status_needs_update_prompt"]) if key in table.index else False)
        assert got == (count, since, prompt), (key, got, (count, since, prompt))
        message = table.loc[key, "status_prompt_message"] if key in table.index else ""
        assert message == (_reference_message(count, since, app_module.norm_label(statuses[key])) if prompt else ""), key


def test_streaks_match_reference_on_sample(app_module, sample_bytes):
    parsed = app_module.parse_workbook_bytes(sample_bytes)
    deck, _ = app_module.build_deck(parsed, AS_OF, False)
    history = parsed["hidden_sheets"]["_Status History"]["table"]
    current = list(zip(deck["deal_key"], deck["status"]))
    _assert_streaks_match(app_module, history, current)
    enriched = app_module.add_status_staleness_columns(deck, app_module.embedded_status_history(parsed), AS_OF)
    assert int(enriched["status_needs_update_prompt"].sum()) == 6


@pytest.mark.parametrize("seed", range(200))
def test_streaks_match_reference_on_random_histories(app_module, seed):
    history, current = _random_history(random.Random(seed))
    _assert_streaks_match(app_module, history, current)


def test_streak_breaks_on_missed_meeting(app_module):
    rows = []
    for day, keys in ((dt.date(2026, 9, 24), ["Bridge::1", "Bridge::2"]),
                      (dt.date(2026, 10, 1), ["Bridge::2"]),
                      (dt.date(2026, 10, 8), ["Bridge::1", "Bridge::2"])):
        for key in keys:
            rows.append({"meeting_date": day.isoformat(), "sheet": "Bridge", "deal_number": key[-1],
                         "deal_key": key, "status": "Surveillance"})
    history = app_module.normalize_status_history(pd.DataFrame(rows))
    table = app_module.status_streak_table(history, (("Bridge::1", "Surveillance"), ("Bridge::2", "Surveillance")), "2026-10-15")
    assert table.loc["Bridge::1", "status_same_meeting_count"] == 2  # 10/01 missing: not 3
    assert table.loc["Bridge::1", "status_same_since"] == "10/08/2026"
    assert table.loc["Bridge::2", "status_same_meeting_count"] == 4


def test_streak_ignores_rows_after_meeting_date(app_module):
    rows = [{"meeting_date": day, "sheet": "Bridge", "deal_number": "10004", "deal_key": "Bridge::10004", "status": status}
            for day, status in (("2026-09-24", "Foreclosure"), ("2026-10-01", "Foreclosure"),
                                ("2026-10-08", "Foreclosure"), ("2026-10-22", "Litigation"))]
    history = app_module.normalize_status_history(pd.DataFrame(rows))
    table = app_module.status_streak_table(history, (("Bridge::10004", "Foreclosure"),), "2026-10-15")
    row = table.loc["Bridge::10004"]
    assert (row["status_same_meeting_count"], row["status_same_since"], bool(row["status_needs_update_prompt"])) == (4, "09/24/2026", True)


def test_streak_edge_cases_do_not_crash(app_module):
    empty = app_module.empty_status_history()
    table = app_module.status_streak_table(empty, (("Bridge::1", ""), ("Bridge::2", "")), "2026-10-15")
    assert table.empty or (table["status_same_meeting_count"] == 0).all()
    deck = pd.DataFrame({"deal_key": ["Bridge::1"], "status": [""]})
    out = app_module.add_status_staleness_columns(deck, empty, AS_OF)
    assert out["status_same_meeting_count"].tolist() == [0]
    assert out["status_same_since"].tolist() == [""]
    assert out["status_needs_update_prompt"].tolist() == [False]
    assert app_module.add_status_staleness_columns(deck.iloc[0:0], empty, AS_OF).empty


# -----------------------------------------------------------------------------
# Loan size: loose money headers, per-row coalescing and the KPI tile (A3)
# -----------------------------------------------------------------------------
@pytest.mark.parametrize(
    "header, expected",
    [("Loan Amount", "loan amount"), ("Orig. Loan Amt ($)", "original loan amount"), ("LoanAmount", "loan amount"),
     ("loan_amount", "loan amount"), ("Loan Amount (USD)", "loan amount"), ("Loan Amount USD", "loan amount"),
     ("Total Commitment ($)", "total commitment"), ("Commit. Amt", "commitment amount"),
     ("Loan Amount ($000s)", "loan amount 000s"), ("Remaining Commitment", "remaining commitment"), (None, "")],
)
def test_loose_header_key(app_module, header, expected):
    assert app_module.loose_header_key(header) == expected


def test_money_roles_match_loosely_and_keep_every_column(app_module):
    headers = ["Deal Number", "Loan Amount", "UPB", "Orig. Loan Amt ($)", "Loan Amount ($000s)", "Total Commitment"]
    layout = app_module.layout_from_header_values(headers, header_row=1)
    assert layout["role_all_labels"]["loan_amount"] == ["Loan Amount", "Orig. Loan Amt ($)"]
    assert layout["role_all_labels"]["commitment"] == ["Total Commitment"]
    assert layout["roles"]["loan_amount"] == "Loan Amount"
    loose_only = app_module.layout_from_header_values(["Deal Number", "Orig Loan Amt"], header_row=1)
    assert loose_only["roles"]["loan_amount"] == "Orig Loan Amt"
    # roles that are not money never match loosely
    assert app_module.layout_from_header_values(["Deal Number", "Status:"], header_row=1)["roles"]["status"] is None


def test_coalesce_amounts_prefers_the_first_positive_value(app_module):
    raw = pd.DataFrame(
        {"a": [None, 0, "$1,000", None, -5], "b": [200, 300, 400, None, None]},
        dtype=object,
    )
    values, headers = app_module._coalesce_amounts(raw, ["a", "b"])
    assert values[:3].tolist() == [200.0, 300.0, 1000.0]
    assert np.isnan(values[3]) and values[4] == -5.0
    assert headers.tolist() == ["b", "b", "a", "", "a"]
    empty_values, empty_headers = app_module._coalesce_amounts(raw, [])
    assert np.isnan(empty_values).all() and set(empty_headers) == {""}


def _kpi_row(**fields):
    base = {"sheet": "Term", "upb": 21_780_000.0, "commitment": 22_000_000.0, "commitment_source": "Original loan amount",
            "commitment_header": "Loan Amount", "maturity_date": pd.Timestamp("2030-11-23"), "days_to_maturity": 1500,
            "next_payment_date": pd.Timestamp("2026-10-30"), "days_to_next_payment": 15, "loan_mod_count": 0,
            "loan_mod_active_count": 0, "loan_mod_latest_type": ""}
    base.update(fields)
    return pd.Series(base)


def test_presentation_kpis(app_module):
    tiles = app_module.build_presentation_kpis(_kpi_row())
    assert [tile[0] for tile in tiles] == ["UPB", "Loan Commitment", "Maturity", "Next Payment", "Loan Mods"]
    assert tiles[1] == ("Loan Commitment", "$22.0MM", "Original loan amount", "rt-signal-muted",
                        "From the Term column 'Loan Amount'")
    assert tiles[2][2:4] == ("In 1,500d", "rt-signal-green")
    assert tiles[4][1:3] == ("0", "No mods")
    missing = app_module.loan_size_kpi(_kpi_row(commitment=np.nan, commitment_source="", commitment_header=""))
    assert missing[1:3] == ("-", "Not in workbook")
    bridge = app_module.loan_size_kpi(_kpi_row(sheet="Bridge", commitment=12_500_000.0, commitment_source="Total commitment",
                                               commitment_header="Loan Commitment"))
    assert bridge[:3] == ("Loan Commitment", "$12.5MM", "Total commitment")
    soon = app_module.build_presentation_kpis(_kpi_row(days_to_maturity=12))
    assert soon[2][2:4] == ("In 12d", "rt-signal-amber")


def test_capital_items_reconcile_in_exact_dollars(app_module):
    bridge = pd.Series({"sheet": "Bridge", "upb": 10_250_000.0, "funded_amount": 10_250_000.0, "commitment": 12_500_000.0,
                        "remaining_commitment": 2_250_000.0})
    items = dict(app_module.build_capital_items(bridge))
    assert (items["Loan Commitment"], items["Funded Amount"], items["Remaining Commitment"]) == (
        "$12,500,000", "$10,250,000", "$2,250,000")
    term = dict(app_module.build_capital_items(_kpi_row(loan_amount=22_000_000.0)))
    assert "Funded Amount" not in term and term["Loan Commitment"] == term["Loan Amount"] == "$22,000,000"


@pytest.mark.parametrize(
    "value, expected",
    [(dt.datetime(2026, 10, 1), "10/01/2026"), (pd.Timestamp("2026-10-01 13:00"), "10/01/2026"), ("90+", "90+"),
     (None, ""), (float("nan"), ""), (pd.NaT, ""), (30, "30"), (" 60+ ", "60+")],
)
def test_fmt_cell_text(app_module, value, expected):
    assert app_module.fmt_cell_text(value) == expected


def test_text_helpers_are_idempotent_on_the_sample(app_module, sample_bytes):
    """A no-op edit compares equal: cleaning loaded text again changes nothing."""
    deck, _ = app_module.build_deck(app_module.parse_workbook_bytes(sample_bytes), AS_OF, True)
    for value in deck["commentary"]:
        assert app_module.clean_free_text(value) == value
    for column in ("status", "owner", "deal_name", "borrower"):
        for value in deck[column]:
            assert app_module.norm_label(value) == value, (column, value)
    for value in deck["deal_number"]:
        assert app_module.norm_id(value) == value
