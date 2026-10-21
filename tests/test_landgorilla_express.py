"""Land Gorilla integration tests. Fully mocked: no network, no real credentials.

The fake backend routes on the same paths the production draw tracker uses:
/api/token, /api/clm/loan?fileNumber=..., /api/clm/pipelineReportTemplates/<id>/loans/<id>,
/api/clm/loan/<id>.
"""
from __future__ import annotations

import datetime as dt
import importlib.util
import io
import os
import sys
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import pytest
import streamlit as st
from openpyxl import Workbook

REPO = Path(__file__).resolve().parents[1]
APP_PATH = REPO / "portfolio_meeting_app.py"
FAKE_PASSWORD = "not-a-real-password-123"
SECRETS = {"landgorilla": {"user": "tester@example.com", "password": FAKE_PASSWORD, "verify": True}}


def _load_app():
    spec = importlib.util.spec_from_file_location("wpo_app_lg_tests", REPO / "landgorilla_live.py")
    module = importlib.util.module_from_spec(spec)
    sys.dont_write_bytecode = True
    sys.modules[spec.name] = module  # dataclasses resolve annotations through sys.modules
    spec.loader.exec_module(module)
    return module


app = _load_app()
TODAY = app.lg_today()


def iso(days: int) -> str:
    return (TODAY + dt.timedelta(days=days)).isoformat()


# -----------------------------------------------------------------------------
# Fake Land Gorilla backend
# -----------------------------------------------------------------------------
class FakeResponse:
    def __init__(self, status: int, payload=None):
        self.status_code = status
        self._payload = payload
        self.ok = 200 <= status < 300

    def json(self):
        if self._payload is None:
            raise ValueError("no JSON")
        return self._payload

    def raise_for_status(self):
        if not self.ok:
            import requests

            raise requests.HTTPError(f"{self.status_code} for url (contains {FAKE_PASSWORD}?)", response=self)


class FakeLG:
    """Routes requests like clmapi.landgorilla.com. Records every call."""

    def __init__(self):
        self.calls = []
        self.loan_lookups = []
        self.token_status = 200
        self.tokens_issued = 0
        self.expired_tokens = set()
        self.loans = {
            "rb010001": 501,  # on track
            "rb010003": 503,  # behind schedule, due date and balance disagree with the workbook
        }
        self.templates = {
            501: {
                "projectCompletedPercentage": 62.5,
                "projectDurationPercentage": 70,
                "loanFundedDate": iso(-300),
                "lastDrawDate": iso(-19),
                "currentLoanDueDate": iso(45),
                "originalLoanDueDate": iso(-47),
                "loanProgram": "Ground Up",
                "riskLabel": [],
                "propertyAddress": "100 Harbor Way",
                "city": "Oakland",
                "state": "CA",
                "borrower": {"borrower": "Harbor View LLC"},
                "dateTimeLastNote": f"{iso(-3)}T15:30:00Z",
                "loanStatus": [{"status": "In Construction"}],
                "systemUsers": ["Pat Lee", "Sam Ortiz"],
            },
            503: {
                "projectCompletedPercentage": "40%",
                "projectDurationPercentage": 0.75,
                "loanFundedDate": iso(-400),
                "lastDrawDate": iso(-120),
                "currentLoanDueDate": iso(-5),
                "originalLoanDueDate": iso(-5),
                "loanProgram": "Fix and Flip",
                "riskLabel": ["Over budget", ""],
                "loanStatus": [{"status": "Active"}],
                "dateTimeLastNote": iso(-75),
                "systemUsers": [],
            },
        }
        self.summaries = {
            501: {"loanBalance": 10_250_000, "balanceToFinish": "1,400,000.00", "lastApprovedDrawEffectiveDate": f"{iso(-20)}T00:00:00"},
            503: {"loanBalance": "8,100,000", "balanceToFinish": 500_000, "lastApprovedDrawEffectiveDate": None},
        }

    def get(self, url, headers=None, data=None, params=None, timeout=None, verify=None):
        assert timeout, "every Land Gorilla call needs a timeout"
        parsed = urlparse(url)
        path, query = parsed.path, parse_qs(parsed.query)
        self.calls.append(path)
        if path == "/api/token":
            assert headers["USER"] == "tester@example.com" and headers["PASSWORD"] == FAKE_PASSWORD
            assert data == {"api_name": "clm"}
            if self.token_status != 200:
                return FakeResponse(self.token_status, {"error": "bad credentials"})
            self.tokens_issued += 1
            expires = (dt.datetime.now(dt.timezone.utc) + dt.timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S")
            return FakeResponse(200, {"token": f"tok-{self.tokens_issued}", "expired": expires})

        token = (headers or {}).get("Authorization", "").removeprefix("Bearer ")
        if not token or token in self.expired_tokens:
            return FakeResponse(401, {"error": "expired"})
        if path == "/api/clm/loan":
            self.loan_lookups.append(query.get("fileNumber", [""])[0])
            loan_id = self.loans.get(query.get("fileNumber", [""])[0])
            return FakeResponse(200, {"data": {"items": [{"id": loan_id}] if loan_id else []}})
        if path.startswith("/api/clm/pipelineReportTemplates/437/loans/"):
            loan_id = int(path.rsplit("/", 1)[-1])
            return FakeResponse(200, {"data": {"fields": self.templates[loan_id]}})
        if path.startswith("/api/clm/loan/"):
            loan_id = int(path.rsplit("/", 1)[-1])
            return FakeResponse(200, {"data": self.summaries[loan_id]})
        return FakeResponse(404, None)


@pytest.fixture(autouse=True)
def clear_caches():
    st.cache_data.clear()
    st.cache_resource.clear()
    yield
    st.cache_data.clear()
    st.cache_resource.clear()


@pytest.fixture
def fake_lg(monkeypatch):
    backend = FakeLG()
    import requests

    monkeypatch.setattr(requests.Session, "get", lambda self, url, **kw: backend.get(url, **kw))
    return backend


def settings():
    return app.lg_settings(SECRETS["landgorilla"])


# -----------------------------------------------------------------------------
# Settings, file numbers, parsing
# -----------------------------------------------------------------------------
def test_settings_absent_or_incomplete_returns_none():
    assert app.lg_settings({}) is None
    assert app.lg_settings({"user": "a"}) is None
    assert app.lg_settings({"password": "b"}) is None


def test_settings_parse_and_never_show_password():
    cfg = app.lg_settings({"user": " a@b.com ", "password": "pw-secret", "verify": "false", "template_id": 512, "file_prefix": "RB0"})
    assert (cfg.user, cfg.verify, cfg.template_id, cfg.file_prefix) == ("a@b.com", False, "512", "rb0")
    assert "pw-secret" not in repr(cfg) and "pw-secret" not in str(cfg)
    assert "pw-secret" not in repr(app.LGClient(cfg, session=FakeLG()))
    assert app.lg_settings({"user": "a", "password": "b"}).verify is True
    assert app.lg_settings({"user": "a", "password": "b"}).template_id == "437"


@pytest.mark.parametrize(
    "value, expected",
    [("10001", "rb010001"), (10001.0, "rb010001"), ("10001.0", "rb010001"), (" 10001 ", "rb010001"),
     ("RB012345", "rb012345"), ("rb0777", "rb0777"), ("", None), (None, None), ("N/A", None), ("12-34", None)],
)
def test_file_number(value, expected):
    assert app.lg_file_number(value, "rb0") == expected


def test_dates_numbers_and_percents():
    assert app._lg_date("2026-10-15T02:30:00Z") == dt.date(2026, 10, 14)  # 10:30 PM Eastern the day before
    assert app._lg_date("2026-10-15T00:00:00") == dt.date(2026, 10, 15)
    assert app._lg_date("10/15/2026") == dt.date(2026, 10, 15)
    assert app._lg_date("garbage") is None and app._lg_date(None) is None and app._lg_date("") is None
    assert app.lg_number("$1,234.50") == 1234.5 and app.lg_number("(1,000)") == -1000 and app.lg_number("x") is None
    assert app.lg_number(float("nan")) is None and app.lg_number(True) is None
    assert app.lg_pct(62.5) == 62.5 and app.lg_pct("40%") == 40 and app.lg_pct(0.75) == 75 and app.lg_pct(150) == 100
    assert app.lg_pct(None) is None and app.lg_pct(0) == 0


def test_parse_template_tolerates_bad_payloads():
    for payload in (None, [], "x", {"data": None}, {"data": {"fields": None}}, {"fields": []}):
        parsed = app.lg_parse_template_loan(payload)
        assert parsed["risk"] == [] and parsed["users"] == [] and parsed["due_date"] is None
    parsed = app.lg_parse_template_loan({"data": {"fields": {"riskLabel": ["A", None, ""], "loanStatus": [{"status": "Active"}], "borrower": "Name"}}})
    assert parsed["risk"] == ["A"] and parsed["status"] == "Active" and parsed["borrower"] == "Name"


def test_md_escape_blocks_latex_and_markup():
    assert app.md_escape("$1 vs $2") == r"\$1 vs \$2"
    assert app.md_escape("a_b*c") == r"a\_b\*c"


# -----------------------------------------------------------------------------
# Fetching and caching
# -----------------------------------------------------------------------------
def test_fetch_ok_not_found_and_one_token(fake_lg):
    results = app.lg_fetch_loans(settings(), ["rb010001", "rb010002", "rb010003", None, "rb010001"])
    assert results["rb010001"]["state"] == "ok" and results["rb010003"]["state"] == "ok"
    assert results["rb010002"]["state"] == "not_found"
    assert fake_lg.tokens_issued == 1
    overview = results["rb010001"]["overview"]
    assert overview["lg_loan_id"] == 501 and overview["status"] == "In Construction" and overview["to_finish"] == "1,400,000.00"
    assert overview["last_approved_draw"] == TODAY - dt.timedelta(days=20)


def test_fetch_uses_cache_until_forced(fake_lg):
    app.lg_fetch_loans(settings(), ["rb010001", "rb010002"])
    calls = len(fake_lg.calls)
    again = app.lg_fetch_loans(settings(), ["rb010001", "rb010002"])
    assert len(fake_lg.calls) == calls and again["rb010001"]["state"] == "ok"
    app.lg_fetch_loans(settings(), ["rb010001"], force=True)
    assert len(fake_lg.calls) > calls


def test_errors_back_off_but_explicit_refresh_retries(fake_lg, monkeypatch):
    original = fake_lg.get

    def flaky(url, **kw):
        if "/api/clm/loan" in url and "fileNumber" in url:
            raise TimeoutError("read timed out")
        return original(url, **kw)

    monkeypatch.setattr(fake_lg, "get", flaky)
    first = app.lg_fetch_loans(settings(), ["rb010001"])
    assert first["rb010001"]["state"] == "error" and first["rb010001"]["error"] == "TimeoutError"
    monkeypatch.setattr(fake_lg, "get", original)
    assert app.lg_fetch_loans(settings(), ["rb010001"])["rb010001"]["state"] == "error"
    assert app.lg_fetch_loans(settings(), ["rb010001"], force=True)["rb010001"]["state"] == "ok"


def test_login_failure_backs_off_and_hides_secret(fake_lg):
    fake_lg.token_status = 401
    first = app.lg_fetch_loans(settings(), ["rb010001"])
    assert first["rb010001"]["state"] == "error" and first["rb010001"]["error"] == "HTTPError 401"
    assert FAKE_PASSWORD not in first["rb010001"]["error"]
    token_calls = fake_lg.calls.count("/api/token")
    app.lg_fetch_loans(settings(), ["rb010001"])  # inside the back-off window: no new login attempt
    assert fake_lg.calls.count("/api/token") == token_calls
    fake_lg.token_status = 200
    assert app.lg_fetch_loans(settings(), ["rb010001"], force=True)["rb010001"]["state"] == "ok"


def test_expired_token_is_refreshed_once(fake_lg):
    client = app.LGClient(settings(), session=fake_lg)
    old = client.token()
    fake_lg.expired_tokens.add(old)
    status, payload = client.get_json("/api/clm/loan/501")
    assert status == 200 and payload["data"]["loanBalance"] == 10_250_000
    assert fake_lg.tokens_issued == 2


def test_many_loans_fetch_in_parallel_with_one_login(fake_lg):
    numbers = [f"rb0{n}" for n in range(20000, 20100)]
    results = app.lg_fetch_loans(settings(), numbers)
    assert len(results) == 100 and all(r["state"] == "not_found" for r in results.values())
    assert fake_lg.tokens_issued == 1


# -----------------------------------------------------------------------------
# Checks: Land Gorilla vs workbook
# -----------------------------------------------------------------------------
def overview_for(loan_id: int) -> dict:
    overview = {"lg_loan_id": loan_id, **app.lg_parse_template_loan({"data": {"fields": FakeLG().templates[loan_id]}})}
    summary = FakeLG().summaries[loan_id]
    overview.update(balance=summary["loanBalance"], to_finish=summary["balanceToFinish"])
    return overview


def test_checks_quiet_when_land_gorilla_and_workbook_agree():
    overview = overview_for(501)
    row = {"upb": 10_250_000, "maturity_date": TODAY + dt.timedelta(days=45)}
    checks = app.lg_checks(row, overview, app.lg_derived(overview, TODAY), TODAY)
    codes = [msg for level, msg in checks if level in ("alert", "warn")]
    assert codes == []
    assert any("Extended 92 days" in msg for _, msg in checks)


def test_checks_flag_every_disagreement():
    overview = overview_for(503)
    row = {"upb": 7_000_000, "maturity_date": TODAY + dt.timedelta(days=60)}
    derived = app.lg_derived(overview, TODAY)
    assert derived["project_pct"] == 40 and derived["duration_pct"] == 75 and derived["gap_pts"] == 35
    checks = app.lg_checks(row, overview, derived, TODAY)
    text = " | ".join(msg for _, msg in checks)
    assert checks[0][0] == "alert" and "Past the Land Gorilla due date" in checks[0][1]
    for expected in ("vs workbook maturity", "Land Gorilla balance $8,100,000 vs workbook UPB $7,000,000",
                     "Construction 40% complete with 75% of the schedule elapsed", "No draw in 120 days",
                     "Last Land Gorilla note was 75 days ago"):
        assert expected in text, expected
    assert "nan" not in text.lower() and "none" not in text.lower()


def test_closed_loan_with_balance_is_flagged():
    overview = {"status": "Paid Off", "due_date": TODAY - dt.timedelta(days=10)}
    checks = app.lg_checks({"upb": 500_000}, overview, app.lg_derived(overview, TODAY), TODAY)
    text = " | ".join(msg for _, msg in checks)
    assert "status is 'Paid Off'" in text and "Past the Land Gorilla due date" not in text


def test_deal_file_numbers_follow_the_chosen_column():
    import pandas as pd

    deck = pd.DataFrame(
        {
            "sheet": ["Bridge", "Bridge", "Term"],
            "deal_key": ["Bridge::1", "Bridge::2", "Term::3"],
            "deal_number": ["10001", "10002", "20001"],
            "Loan Number": ["555", "", "777"],
        }
    )
    assert app.lg_deal_file_numbers(deck, "deal_number", "rb0") == {"Bridge::1": "rb010001", "Bridge::2": "rb010002"}
    assert app.lg_deal_file_numbers(deck, "Loan Number", "rb0") == {"Bridge::1": "rb0555", "Bridge::2": None}
    assert "Loan Number" in app.lg_loan_column_options(deck)


# -----------------------------------------------------------------------------
# End to end with AppTest
# -----------------------------------------------------------------------------
def build_workbook_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Bridge"
    ws.append(["Portfolio Overview - Bridge"])
    ws.append([
        "Deal Number", "Deal Name", "Borrower Name", "Point Person", "Status", "AM Commentary",
        "Loan Commitment", "Active Funded Amount", "Remaining Commitment", "UPB",
        "Current Maturity Date", "Next Payment Date", "Days Past Due",
    ])
    ws.append([10001, "Harbor View Lofts", "Harbor View LLC", "Alice Chen", "Surveillance", "On track.",
               12_500_000, 10_250_000, 2_250_000, 10_250_000, TODAY + dt.timedelta(days=45), TODAY + dt.timedelta(days=5), 0])
    ws.append(["10002", "Maple Street Rehab", "Maple St Partners", "Bob Diaz", "Upcoming Maturity", "Extension pending.",
               3_400_000, 3_400_000, 0, 3_400_000, TODAY + dt.timedelta(days=12), TODAY - dt.timedelta(days=3), 3])
    ws.append(["10003", "Sunset Ridge Townhomes", "Sunset Ridge Dev Co", "Alice Chen", "Modification", "Mod approved.",
               8_750_000, 7_000_000, 1_750_000, 7_000_000, TODAY + dt.timedelta(days=60), TODAY - dt.timedelta(days=15), 15])
    term = wb.create_sheet("Term")
    term.append(["Deal Number", "Loan Name", "Asset Manager", "Status", "AM Commentary", "Loan Amount", "UPB",
                 "Maturity Date", "Next Payment Date", "Days Past Due"])
    term.append(["20001", "Bayside Rental Portfolio", "Erin Fox", "Surveillance", "Performing.", 22_000_000, 21_780_000,
                 TODAY + dt.timedelta(days=1500), TODAY + dt.timedelta(days=15), 0])
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def make_app_test(tmp_path, monkeypatch, *, with_secrets=True):
    from streamlit.testing.v1 import AppTest

    (tmp_path / "portfolio_overview.xlsx").write_bytes(build_workbook_bytes())
    monkeypatch.chdir(tmp_path)
    at = AppTest.from_file(str(APP_PATH), default_timeout=120)
    if with_secrets:
        at.secrets["landgorilla"] = dict(SECRETS["landgorilla"])
    return at


def all_text(at) -> str:
    parts = []
    for kind in ("markdown", "caption", "info", "warning", "error", "success", "subheader", "metric"):
        for element in getattr(at, kind):
            parts.append(str(getattr(element, "value", "")))
            parts.append(str(getattr(element, "label", "")))
    for element in at.sidebar.markdown:
        parts.append(str(element.value))
    return " ".join(parts)


def load_construction(at):
    next(button for button in at.button if button.label == "Load construction data").click().run()
    assert not at.exception


def test_bridge_slide_loads_only_selected_loan_on_request(tmp_path, monkeypatch, fake_lg):
    at = make_app_test(tmp_path, monkeypatch).run()
    assert not at.exception
    assert fake_lg.calls == []
    panel = next(exp for exp in at.expander if exp.label == "Construction · Land Gorilla")
    assert not panel.proto.expanded
    load_construction(at)
    text = all_text(at)
    assert "File rb010001" in text
    assert "Construction 62% complete" in " ".join(str(p.proto.text) for p in at.get("progress"))
    assert {"LG balance", "Construction holdback remaining", "Last approved draw", "LG due date"} <= {m.label for m in at.metric}
    assert fake_lg.loan_lookups == ["rb010001"]
    assert FAKE_PASSWORD not in text


def test_bridge_slide_shows_neutral_metadata_and_source_differences(tmp_path, monkeypatch, fake_lg):
    fake_lg.templates[503]["riskLabel"] = ["GUC", "10% Retainage", "PERMITS REQUIRED"]
    at = make_app_test(tmp_path, monkeypatch).run()
    jump_lg(at, "Sunset")
    assert fake_lg.calls == []
    load_construction(at)
    captions = " ".join(str(c.value) for c in at.caption)
    assert r"LG balance \$8,100,000; workbook UPB \$7,000,000" in captions
    assert "LG labels / instructions: GUC" in captions
    assert "10% Retainage" in captions
    assert not at.error and not at.warning
    jump_lg(at, "Maple")
    load_construction(at)
    assert "No matching Land Gorilla file was returned" in all_text(at)
    assert not at.exception


def test_term_slide_and_overview_use_snapshots_only(tmp_path, monkeypatch, fake_lg):
    at = make_app_test(tmp_path, monkeypatch).run()
    load_construction(at)
    jump_lg(at, "Sunset")
    load_construction(at)
    count = len(fake_lg.calls)
    jump_lg(at, "Bayside")
    assert not any(b.label.endswith("construction data") for b in at.button)
    at.sidebar.radio[0].set_value("Overview").run()
    assert not at.exception
    assert len(fake_lg.calls) == count
    frames = [df.value for df in at.dataframe]
    lg_frame = next(f for f in frames if "Built" in f.columns)
    assert set(lg_frame["Deal #"]) == {"10003", "10001"}


def test_configured_api_never_runs_during_navigation_or_editing(tmp_path, monkeypatch, fake_lg):
    # Empty, expired and failed caches must not trigger automatic retries anywhere.
    import landgorilla_live as live
    at = make_app_test(tmp_path, monkeypatch).run()
    assert not at.exception and not fake_lg.calls
    for loan in ("Maple", "Sunset", "Bayside", "Harbor"):
        jump_lg(at, loan)
        assert not at.exception and not fake_lg.calls
    cfg = live.lg_settings(SECRETS["landgorilla"])
    store = live._lg_account_store(cfg)
    store["loans"]["rb010001"] = {"state": "ok", "fetched_at": 1, "overview": overview_for(501)}
    at.run()
    assert "Saved snapshot" in all_text(at) and not fake_lg.calls
    store.setdefault("read_failures", {})["rb010001"] = {
        "state": "error", "fetched_at": 1, "checked_at": 1, "overview": None, "error": "HTTP 500"}
    for view in ("Overview", "Review", "Exports", "Presentation"):
        at.sidebar.radio[0].set_value(view).run()
        assert not at.exception and not fake_lg.calls
    next(b for b in at.button if b.label.startswith("Edit status")).click().run()
    assert not at.exception and not fake_lg.calls
    assert set(at.sidebar.radio[0].options) == {"Overview", "Presentation", "Review", "Exports"}
    assert "m" not in at.query_params
    assert not any("Meeting capture" in e.label or "Meeting backup" in e.label for e in at.expander)


def test_not_configured_changes_nothing(tmp_path, monkeypatch, fake_lg):
    at = make_app_test(tmp_path, monkeypatch, with_secrets=False)
    at.run()
    assert not at.exception
    assert "Land Gorilla · live" not in all_text(at)
    assert any("Not configured" in str(c.value) for c in at.sidebar.caption)
    assert fake_lg.calls == []


def test_failing_login_keeps_the_meeting_running(tmp_path, monkeypatch, fake_lg):
    fake_lg.token_status = 401
    at = make_app_test(tmp_path, monkeypatch)
    at.run()
    assert not at.exception
    assert fake_lg.calls == []
    load_construction(at)
    assert "HTTPError 401" in all_text(at)
    assert not at.warning and not at.error
    assert FAKE_PASSWORD not in all_text(at)
    at.sidebar.radio[0].set_value("Overview").run()
    assert not at.exception



def jump_lg(at, name):
    box = next(element for element in at.selectbox if element.label == "Jump to deal")
    box.set_value(next(option for option in box.options if name.casefold() in option.casefold())).run()


def test_cache_isolated_after_settings_change(fake_lg):
    app.lg_fetch_loans(settings(), ["rb010001"])
    count = fake_lg.calls.count("/api/token")
    other = app.LGSettings(user="different-account", password="different-password")
    result = app.lg_fetch_loans(other, ["rb010001"])
    assert fake_lg.calls.count("/api/token") == count + 1
    assert result["rb010001"]["state"] == "error"


@pytest.mark.parametrize("items,state,loan_id", [
    ([{"id": 503, "fileNumber": "rb010003"}, {"id": 501, "fileNumber": "RB010001"}], "ok", 501),
    ([{"id": 503, "fileNumber": "rb010003"}], "not_found", None),
    ([{"id": 501}, {"id": 503}], "error", None),
])
def test_lookup_does_not_select_an_unrelated_or_ambiguous_loan(fake_lg, monkeypatch, items, state, loan_id):
    original = fake_lg.get
    def lookup(url, **kwargs):
        if urlparse(url).path == "/api/clm/loan":
            return FakeResponse(200, {"data": {"items": items}})
        return original(url, **kwargs)
    monkeypatch.setattr(fake_lg, "get", lookup)
    result = app.lg_fetch_loans(settings(), ["rb010001"])["rb010001"]
    assert result["state"] == state
    assert (result.get("overview") or {}).get("lg_loan_id") == loan_id


def test_failed_login_on_refresh_preserves_and_labels_previous_snapshot(fake_lg):
    previous = app.lg_fetch_loans(settings(), ["rb010001"])["rb010001"]
    fake_lg.token_status = 401
    failed = app.lg_fetch_loans(settings(), ["rb010001"], force=True)["rb010001"]
    assert failed["stale"] and failed["refresh_error"] == "HTTPError 401"
    assert failed["fetched_at"] == previous["fetched_at"]
    assert failed["overview"] == previous["overview"]


def test_diagnostic_output_omits_sensitive_values():
    from tools.diagnose_landgorilla import report
    overview = {"balance": 1234567, "borrower": "Private borrower", "file_number": "rb099999", "project_pct": 0}
    result = {"state": "ok", "overview": overview, "error": ""}
    summary = report(result, [{"endpoint": "/api/clm/loan/<id>", "status": 200}])
    assert summary["fields_present"]["balance"] and summary["fields_present"]["project_pct"]
    assert not summary["fields_present"]["due_date"]
    assert "1234567" not in str(summary) and "Private borrower" not in str(summary) and "rb099999" not in str(summary)


def test_login_backoff_isolated_between_accounts(fake_lg):
    fake_lg.token_status = 401
    app.lg_fetch_loans(settings(), ["rb010001"])
    count = fake_lg.calls.count("/api/token")
    app.lg_fetch_loans(app.LGSettings(user="another-account", password="other"), ["rb010001"])
    assert fake_lg.calls.count("/api/token") == count + 1


def test_lookup_http_500_recovers_after_bounded_retries(fake_lg, monkeypatch):
    original, attempts = fake_lg.get, []
    monkeypatch.setattr(app.time, "sleep", lambda seconds: None)
    def flaky(url, **kwargs):
        if urlparse(url).path == "/api/clm/loan":
            attempts.append(url)
            if len(attempts) < 3:
                return FakeResponse(500, {"error": "temporary server failure"})
        return original(url, **kwargs)
    monkeypatch.setattr(fake_lg, "get", flaky)
    result = app.lg_fetch_loans(settings(), ["rb010001"])["rb010001"]
    assert result["state"] == "ok" and result["overview"]["lg_loan_id"] == 501
    assert len(attempts) == 3


def test_persistent_500_keeps_last_success_and_does_not_spam_reruns(fake_lg, monkeypatch):
    previous = app.lg_fetch_loans(settings(), ["rb010001"])["rb010001"]
    original, attempts = fake_lg.get, []
    monkeypatch.setattr(app.time, "sleep", lambda seconds: None)
    def failing(url, **kwargs):
        if urlparse(url).path == "/api/clm/loan":
            attempts.append(url)
            return FakeResponse(500, {"error": FAKE_PASSWORD})
        return original(url, **kwargs)
    monkeypatch.setattr(fake_lg, "get", failing)
    failed = app.lg_fetch_loans(settings(), ["rb010001"], force=True)["rb010001"]
    assert failed["stale"] is True and failed["overview"] == previous["overview"]
    assert failed["fetched_at"] == previous["fetched_at"]
    assert failed["refresh_error"] == "HTTP 500" and FAKE_PASSWORD not in str(failed)
    assert len(attempts) == app.LG_READ_ATTEMPTS
    again = app.lg_fetch_loans(settings(), ["rb010001"])["rb010001"]
    assert again["stale"] and len(attempts) == app.LG_READ_ATTEMPTS


def test_persistent_500_without_previous_data_remains_an_error(fake_lg, monkeypatch):
    original, attempts = fake_lg.get, []
    monkeypatch.setattr(app.time, "sleep", lambda seconds: None)
    def failing(url, **kwargs):
        if urlparse(url).path == "/api/clm/loan":
            attempts.append(url)
            return FakeResponse(500, None)
        return original(url, **kwargs)
    monkeypatch.setattr(fake_lg, "get", failing)
    result = app.lg_fetch_loans(settings(), ["rb010001"])["rb010001"]
    assert result["state"] == "error" and result["overview"] is None
    app.lg_fetch_loans(settings(), ["rb010001"])
    assert len(attempts) == 3


def test_timeout_retries_and_recovers(fake_lg, monkeypatch):
    import requests
    original, attempts = fake_lg.get, []
    monkeypatch.setattr(app.time, "sleep", lambda seconds: None)
    def flaky(url, **kwargs):
        if urlparse(url).path == "/api/clm/loan":
            attempts.append(url)
            if len(attempts) < 3:
                raise requests.Timeout("sensitive response omitted")
        return original(url, **kwargs)
    monkeypatch.setattr(fake_lg, "get", flaky)
    assert app.lg_fetch_loans(settings(), ["rb010001"])["rb010001"]["state"] == "ok"
    assert len(attempts) == 3


def test_both_detail_requests_failing_do_not_cache_empty_success(fake_lg, monkeypatch):
    original = fake_lg.get
    monkeypatch.setattr(app.time, "sleep", lambda seconds: None)
    def failing(url, **kwargs):
        if "/loans/" in url or urlparse(url).path == "/api/clm/loan/501":
            return FakeResponse(500, None)
        return original(url, **kwargs)
    monkeypatch.setattr(fake_lg, "get", failing)
    result = app.lg_fetch_loans(settings(), ["rb010001"])["rb010001"]
    assert result["state"] == "error" and result["overview"] is None


def test_partial_detail_response_is_explicitly_labeled(fake_lg, monkeypatch):
    original = fake_lg.get
    monkeypatch.setattr(app.time, "sleep", lambda seconds: None)
    def failing(url, **kwargs):
        if "/loans/" in url:
            return FakeResponse(500, None)
        return original(url, **kwargs)
    monkeypatch.setattr(fake_lg, "get", failing)
    result = app.lg_fetch_loans(settings(), ["rb010001"])["rb010001"]
    assert result["state"] == "ok" and result["overview"]["balance"] == 10_250_000
    assert result["partial_errors"] == ["Construction details unavailable (HTTP 500)"]


def test_rate_limit_retry_after_is_bounded(fake_lg, monkeypatch):
    original, attempts, delays = fake_lg.get, [], []
    monkeypatch.setattr(app.time, "sleep", delays.append)
    def limited(url, **kwargs):
        if urlparse(url).path == "/api/clm/loan":
            attempts.append(url)
            if len(attempts) == 1:
                response = FakeResponse(429, None)
                response.headers = {"Retry-After": "120"}
                return response
        return original(url, **kwargs)
    monkeypatch.setattr(fake_lg, "get", limited)
    assert app.lg_fetch_loans(settings(), ["rb010001"])["rb010001"]["state"] == "ok"
    assert delays == [2.0]


def test_batch_deadline_stops_queued_requests_without_network_call(fake_lg):
    import requests
    client = app.LGClient(settings(), session=fake_lg)
    client.token()
    count = len(fake_lg.calls)
    client._read_deadline = app.time.monotonic() - 1
    with pytest.raises(requests.Timeout, match="read budget"):
        client.get_json("/api/clm/loan?fileNumber=rb010001")
    assert len(fake_lg.calls) == count


def test_app_http_500_keeps_workbook_usable(tmp_path, monkeypatch, fake_lg):
    original = fake_lg.get
    monkeypatch.setattr(app.time, "sleep", lambda seconds: None)
    def failing(url, **kwargs):
        if urlparse(url).path == "/api/clm/loan":
            return FakeResponse(500, None)
        return original(url, **kwargs)
    monkeypatch.setattr(fake_lg, "get", failing)
    at = make_app_test(tmp_path, monkeypatch).run()
    assert not at.exception
    load_construction(at)
    assert "HTTP 500" in all_text(at)
    assert not at.warning and not at.error
    assert "Harbor View" in all_text(at)
    at.sidebar.radio(key="view_mode").set_value("Overview").run()
    assert not at.exception


def test_app_failed_refresh_labels_last_known_data(tmp_path, monkeypatch, fake_lg):
    at = make_app_test(tmp_path, monkeypatch).run()
    load_construction(at)
    original = fake_lg.get
    monkeypatch.setattr(app.time, "sleep", lambda seconds: None)
    def failing(url, **kwargs):
        if urlparse(url).path == "/api/clm/loan":
            return FakeResponse(500, None)
        return original(url, **kwargs)
    monkeypatch.setattr(fake_lg, "get", failing)
    next(button for button in at.button if button.label == "Refresh construction data").click().run()
    assert not at.exception
    assert "These are the last retrieved values" in all_text(at)
    assert "Refresh failed (HTTP 500)" in all_text(at)
    assert not at.warning and not at.error
