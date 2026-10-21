"""Live Land Gorilla construction data preserved from the deployed integration."""
from __future__ import annotations

import datetime as dt
import hashlib
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Mapping, Optional, Tuple
from urllib.parse import quote, urlencode
import numpy as np
import pandas as pd
import streamlit as st


def norm_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    return re.sub(r"^(\d+)\.0+$", r"\1", text)


def display_text(value, blank="-"):
    return norm_text(value) or blank


def first_nonblank_value(*values):
    return next((value for value in values if norm_text(value)), None)


def safe_series(frame, column, default=""):
    return frame[column] if column in frame else pd.Series(default, index=frame.index)


def coerce_datetime_value(value):
    if value is None or isinstance(value, (bool, int, float, np.number)):
        return pd.NaT
    try:
        stamp = pd.to_datetime(value, errors="coerce", utc=True)
        return stamp.tz_convert(None) if pd.notna(stamp) else pd.NaT
    except (TypeError, ValueError, OverflowError):
        return pd.NaT


def fmt_date(value):
    stamp = coerce_datetime_value(value)
    return "-" if pd.isna(stamp) else stamp.strftime("%m/%d/%Y")


# -----------------------------------------------------------------------------
# Land Gorilla (live construction data for Bridge loans)
# -----------------------------------------------------------------------------
# Land Gorilla (LG) is the construction-loan system of record, so for Bridge loans it is
# fresher than the weekly workbook. The calls mirror the production draw tracker: a token
# from /api/token (USER/PASSWORD headers), the LG loan id from
# /api/clm/loan?fileNumber=rb0<loan #>, then the pipeline-template fields and the loan
# summary. Credentials come only from st.secrets["landgorilla"]. Exception text is never
# shown, because it could echo a malformed secret to every viewer.
LG_BASE_URL = "https://clmapi.landgorilla.com"
LG_BASE_URL_ENV = "WPO_LG_BASE_URL"  # local fake servers only (tests / screenshots)
LG_TEMPLATE_ID_DEFAULT = "437"
LG_FILE_PREFIX_DEFAULT = "rb0"
LG_TIMEZONE = "America/New_York"
LG_CACHE_SECONDS = 30 * 60
LG_NOT_FOUND_SECONDS = 30 * 60
LG_LOGIN_BACKOFF_SECONDS = 10 * 60  # a failing login is not retried on every rerun (account lockouts)
LG_MAX_WORKERS = 4
LG_READ_ATTEMPTS = 3
LG_ERROR_RETRY_SECONDS = 60
LG_BATCH_READ_SECONDS = 30
LG_RETRY_STATUS = (429, 500, 502, 503, 504)
LG_TIMEOUT_SECONDS = 30
LG_STORE_MAX_LOANS = 5000
LG_DEFAULT_LOAN_COLUMN = "deal_number"

# Thresholds for the Land Gorilla vs. workbook checks.
LG_BEHIND_SCHEDULE_POINTS = 20
LG_STALLED_DRAW_DAYS = 90
LG_STALE_NOTE_DAYS = 60
LG_DUE_DATE_TOLERANCE_DAYS = 3
LG_BALANCE_TOLERANCE = 1_000
LG_BALANCE_TOLERANCE_PCT = 0.01
LG_CLOSED_STATUS_WORDS = ("paid", "closed", "payoff", "pay off")


@dataclass(frozen=True)
class LGSettings:
    user: str
    password: str = field(repr=False)
    verify: bool = True
    template_id: str = LG_TEMPLATE_ID_DEFAULT
    file_prefix: str = LG_FILE_PREFIX_DEFAULT


@dataclass
class LGContext:
    settings: LGSettings
    loan_column: str
    deal_files: Dict[str, Optional[str]]  # Bridge deal_key -> LG file number (None when unusable)
    results: Dict[str, dict]  # file number -> fetch result


def _lg_truthy(value: object, default: bool = True) -> bool:
    if isinstance(value, bool):
        return value
    text = norm_text(value).lower()
    if not text:
        return default
    return text not in {"false", "0", "no", "off", "n"}


def lg_settings(secrets: Optional[Mapping[str, object]] = None) -> Optional[LGSettings]:
    """Land Gorilla settings from st.secrets["landgorilla"]; None when not configured."""
    try:
        if secrets is None:
            section = st.secrets.get("landgorilla", None)
            secrets = dict(section) if section else {}
        else:
            secrets = dict(secrets)
    except Exception:  # no secrets.toml at all
        return None

    user = str(secrets.get("user") or "").strip()
    password = str(secrets.get("password") or "")
    if not user or not password:
        return None

    template_id = norm_text(
        first_nonblank_value(
            secrets.get("template_id"),
            secrets.get("pipeline_template_id"),
            secrets.get("template"),
        )
    )
    prefix = norm_text(first_nonblank_value(secrets.get("file_prefix"), secrets.get("file_number_prefix")))
    return LGSettings(
        user=user,
        password=password,
        verify=_lg_truthy(secrets.get("verify"), default=True),
        template_id=template_id or LG_TEMPLATE_ID_DEFAULT,
        file_prefix=(prefix or LG_FILE_PREFIX_DEFAULT).lower(),
    )


def lg_base_url() -> str:
    override = os.environ.get(LG_BASE_URL_ENV, "").strip().rstrip("/")
    if override and re.fullmatch(r"http://(127\.0\.0\.1|localhost)(:\d+)?", override):
        return override
    return LG_BASE_URL


def lg_error_text(exc: BaseException) -> str:
    """Exception type (plus HTTP status) only; never the message, which could contain a secret."""
    status = getattr(getattr(exc, "response", None), "status_code", None)
    name = type(exc).__name__
    return f"{name} {status}" if status else name


def lg_file_number(value: object, prefix: str = LG_FILE_PREFIX_DEFAULT) -> Optional[str]:
    """LG file number for a workbook loan number: rb0 + digits, or an already-prefixed value."""
    text = norm_text(value).replace(" ", "").lower()
    if not text:
        return None
    if text.startswith(prefix.lower()) or re.fullmatch(r"rb\d+", text):
        return text
    if re.fullmatch(r"\d+", text):
        return f"{prefix.lower()}{text}"
    return None


def lg_today() -> dt.date:
    return pd.Timestamp.now(tz=LG_TIMEZONE).date()


def _lg_date(value: object) -> Optional[dt.date]:
    """An LG date or timestamp as a date; zone-aware timestamps are read in Eastern time."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = str(value).strip()
    if not text:
        return None
    if "T" in text and re.search(r"(?:Z|[+-]\d{2}:?\d{2})$", text):
        stamp = pd.to_datetime(text, errors="coerce", utc=True)
        if pd.notna(stamp):
            return stamp.tz_convert(LG_TIMEZONE).date()
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(text[:19], fmt).date()
        except ValueError:
            continue
    stamp = pd.to_datetime(text, errors="coerce")
    return None if pd.isna(stamp) else pd.Timestamp(stamp).date()


def lg_items_of(payload: object) -> list:
    data = payload.get("data", payload) if isinstance(payload, dict) else payload
    if isinstance(data, dict):
        items = data.get("items", [])
        return items if isinstance(items, list) else []
    return data if isinstance(data, list) else []


def lg_parse_template_loan(payload: object) -> dict:
    """Loan-level construction progress from a pipeline-template loan payload."""
    fields = {}
    if isinstance(payload, dict):
        data = payload.get("data")
        fields = (data.get("fields") if isinstance(data, dict) else None) or payload.get("fields") or {}
    if not isinstance(fields, dict):
        fields = {}

    borrower = fields.get("borrower") or {}
    statuses = fields.get("loanStatus") or []
    first_status = statuses[0] if isinstance(statuses, list) and statuses else {}
    risk = fields.get("riskLabel") or []
    users = fields.get("systemUsers") or []
    return {
        "project_pct": fields.get("projectCompletedPercentage"),
        "duration_pct": fields.get("projectDurationPercentage"),
        "funded_date": _lg_date(fields.get("loanFundedDate")),
        "last_draw": _lg_date(fields.get("lastDrawDate")),
        "due_date": _lg_date(fields.get("currentLoanDueDate")),
        "original_due": _lg_date(fields.get("originalLoanDueDate")),
        "program": norm_text(fields.get("loanProgram")),
        "risk": [norm_text(x) for x in risk if norm_text(x)] if isinstance(risk, list) else [],
        "property": norm_text(fields.get("propertyAddress")),
        "city": norm_text(fields.get("city")),
        "state": norm_text(fields.get("state")),
        "borrower": norm_text(borrower.get("borrower") if isinstance(borrower, dict) else borrower),
        "last_note_at": _lg_date(fields.get("dateTimeLastNote")),
        "status": norm_text(first_status.get("status") if isinstance(first_status, dict) else first_status),
        "users": [norm_text(u) for u in users if norm_text(u)] if isinstance(users, list) else [],
    }


def lg_number(value: object) -> Optional[float]:
    """Numbers or strings like "$1,234.50" / "(1,000)" as floats."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float, np.integer, np.floating)):
        number = float(value)
        return None if np.isnan(number) or np.isinf(number) else number
    text = str(value).strip()
    negative = text.startswith("(") and text.endswith(")")
    text = re.sub(r"[,$\s()%]", "", text)
    try:
        number = float(text)
    except ValueError:
        return None
    if np.isnan(number) or np.isinf(number):
        return None
    return -number if negative else number


def lg_pct(value: object) -> Optional[float]:
    """A percent on a 0-100 scale. LG reports 0-100; a value strictly between 0 and 1 is read
    as a fraction, because "0.62% built" is far less likely than 62%."""
    number = lg_number(value)
    if number is None:
        return None
    if 0 < number < 1:
        number *= 100
    return float(min(max(number, 0.0), 100.0))


def lg_derived(overview: Mapping[str, object], today: dt.date) -> Dict[str, object]:
    def days_since(day: object) -> Optional[int]:
        return (today - day).days if isinstance(day, dt.date) else None

    project = lg_pct(overview.get("project_pct"))
    duration = lg_pct(overview.get("duration_pct"))
    due = overview.get("due_date")
    original = overview.get("original_due")
    extension_days = 0
    if isinstance(due, dt.date) and isinstance(original, dt.date) and due > original:
        extension_days = (due - original).days

    return {
        "project_pct": project,
        "duration_pct": duration,
        "gap_pts": (duration - project) if project is not None and duration is not None else None,
        "days_since_draw": days_since(overview.get("last_draw")),
        "days_since_funding": days_since(overview.get("funded_date")),
        "days_since_note": days_since(overview.get("last_note_at")),
        "extension_days": extension_days,
        "days_to_due": (due - today).days if isinstance(due, dt.date) else None,
        "balance": lg_number(overview.get("balance")),
        "to_finish": lg_number(overview.get("to_finish")),
    }


def lg_status_closed(overview: Mapping[str, object]) -> bool:
    status = norm_text(overview.get("status")).lower()
    return any(word in status for word in LG_CLOSED_STATUS_WORDS)


def fmt_money_exact(value: object, blank: str = "-") -> str:
    number = lg_number(value)
    if number is None:
        return blank
    sign = "-" if number < 0 else ""
    return f"{sign}${abs(number):,.0f}"


def lg_checks(row: Mapping[str, object], overview: Mapping[str, object], derived: Mapping[str, object], today: dt.date) -> List[Tuple[str, str]]:
    """(level, message) pairs where Land Gorilla and the workbook disagree or the loan needs a look.
    Levels: alert, warn, info."""
    checks: List[Tuple[str, str]] = []
    closed = lg_status_closed(overview)

    due = overview.get("due_date")
    maturity = coerce_datetime_value(row.get("maturity_date"))
    if isinstance(due, dt.date) and not pd.isna(maturity):
        gap_days = abs((pd.Timestamp(due) - pd.Timestamp(maturity).normalize()).days)
        if gap_days > LG_DUE_DATE_TOLERANCE_DAYS:
            checks.append(("warn", f"Land Gorilla due date {fmt_date(due)} vs workbook maturity {fmt_date(maturity)}."))

    if isinstance(due, dt.date) and due < today and not closed:
        checks.append(("alert", f"Past the Land Gorilla due date ({fmt_date(due)})."))

    extension_days = int(derived.get("extension_days") or 0)
    if extension_days > 0:
        checks.append(("info", f"Extended {extension_days:,} days in Land Gorilla (original due {fmt_date(overview.get('original_due'))})."))

    balance = derived.get("balance")
    upb = lg_number(row.get("upb"))
    if balance is not None and upb is not None:
        tolerance = max(LG_BALANCE_TOLERANCE, abs(upb) * LG_BALANCE_TOLERANCE_PCT)
        if abs(balance - upb) > tolerance:
            checks.append(("warn", f"Land Gorilla balance {fmt_money_exact(balance)} vs workbook UPB {fmt_money_exact(upb)}."))

    project = derived.get("project_pct")
    duration = derived.get("duration_pct")
    gap = derived.get("gap_pts")
    if gap is not None and gap >= LG_BEHIND_SCHEDULE_POINTS:
        checks.append(("warn", f"Construction {project:.0f}% complete with {duration:.0f}% of the schedule elapsed."))

    if project is not None and project < 100 and not closed:
        idle = derived.get("days_since_draw")
        if idle is None:
            idle = derived.get("days_since_funding")
        if idle is not None and idle >= LG_STALLED_DRAW_DAYS:
            checks.append(("warn", f"No draw in {idle:,} days while construction is {project:.0f}% complete."))

    if closed and upb is not None and upb > 0:
        checks.append(("warn", f"Land Gorilla status is '{overview.get('status')}' but the workbook shows UPB {fmt_money_exact(upb)}."))

    note_days = derived.get("days_since_note")
    if note_days is not None and note_days >= LG_STALE_NOTE_DAYS:
        checks.append(("info", f"Last Land Gorilla note was {note_days:,} days ago."))

    order = {"alert": 0, "warn": 1, "info": 2}
    return sorted(checks, key=lambda item: order.get(item[0], 3))


@st.cache_resource(show_spinner=False)
def _lg_install_truststore() -> bool:
    """Use the OS trust store when truststore is installed (corporate TLS); a no-op otherwise."""
    try:
        import truststore

        truststore.inject_into_ssl()
        return True
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def _lg_store() -> dict:
    """Process-wide cache of Land Gorilla reads: {"loans": {file number: result}, ...}.
    Errors are not cached, so a timeout is simply retried on the next run."""
    return {"lock": threading.Lock(), "loans": {}, "login_failure": None}


def _lg_account_store(settings=None):
    """Cache results and failed logins separately for each account and API configuration."""
    root = _lg_store()
    if settings is None:
        settings = lg_settings()
    if settings is None:
        return root
    identity = repr((settings.user, settings.password, settings.verify, settings.template_id,
                     settings.file_prefix, lg_base_url()))
    namespace = hashlib.sha256(identity.encode("utf-8")).hexdigest()
    with root["lock"]:
        accounts = root.setdefault("accounts", {})
        return accounts.setdefault(namespace, {"lock": threading.Lock(), "loans": {}, "login_failure": None})


def _lg_close_sessions(sessions):
    for session in sessions:
        close = getattr(session, "close", None)
        if close:
            try:
                close()
            except Exception:
                pass


def _lg_new_session():
    import requests

    return requests.Session()


class LGClient:
    """Minimal Land Gorilla CLM client: one bearer token, refreshed shortly before it expires."""

    def __init__(self, settings: LGSettings, session=None) -> None:
        self._settings = settings
        self._session = session if session is not None else _lg_new_session()
        self._token: Optional[str] = None
        self._expires = 0.0
        self._lock = threading.Lock()
        self.base_url = lg_base_url()

    def __repr__(self) -> str:
        return f"LGClient(user={self._settings.user!r})"

    def token(self, refresh: bool = False) -> str:
        with self._lock:
            if self._token and not refresh and time.time() < self._expires - 120:
                return self._token
            response = self._session.get(
                f"{self.base_url}/api/token",
                headers={"USER": self._settings.user, "PASSWORD": self._settings.password},
                data={"api_name": "clm"},
                timeout=20,
                verify=self._settings.verify,
            )
            response.raise_for_status()
            payload = response.json()
            self._token = str(payload["token"])
            try:
                expires = dt.datetime.strptime(str(payload["expired"]), "%Y-%m-%d %H:%M:%S")
                self._expires = expires.replace(tzinfo=dt.timezone.utc).timestamp()
            except Exception:
                self._expires = time.time() + 1800
            return self._token

    def _read(self, http, path, token):
        """Retry transient read failures; never retry rejected credentials here."""
        import requests

        for attempt in range(LG_READ_ATTEMPTS):
            response = None
            deadline = getattr(self, "_read_deadline", None)
            remaining = deadline - time.monotonic() if deadline is not None else LG_TIMEOUT_SECONDS
            if remaining <= 0:
                raise requests.Timeout("Land Gorilla read budget exhausted")
            try:
                response = http.get(
                    f"{self.base_url}{path}",
                    headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
                    timeout=min(LG_TIMEOUT_SECONDS, max(0.1, remaining)), verify=self._settings.verify,
                )
                if response.status_code not in LG_RETRY_STATUS or attempt == LG_READ_ATTEMPTS - 1:
                    return response
            except (requests.Timeout, requests.ConnectionError):
                if attempt == LG_READ_ATTEMPTS - 1:
                    raise
            delay = 0.4 * (2 ** attempt)
            try:
                retry_after = float((getattr(response, "headers", {}) or {}).get("Retry-After", 0))
                delay = max(delay, min(retry_after, 2.0))
            except (TypeError, ValueError):
                pass
            if deadline is not None:
                delay = min(delay, max(0.0, deadline - time.monotonic()))
            time.sleep(delay)

    def get_json(self, path: str, session=None) -> Tuple[int, object]:
        """GET a CLM path: (status code, parsed JSON or None). A 401 refreshes the token once."""
        http = session if session is not None else self._session
        status = 0
        for attempt in range(2):
            token = self.token(refresh=attempt > 0)
            response = self._read(http, path, token)
            status = int(response.status_code)
            if status == 401 and attempt == 0:
                continue
            if not response.ok:
                return status, None
            try:
                return status, response.json()
            except ValueError:
                return status, None
        return status, None


def _lg_result(state: str, overview: Optional[dict] = None, error: str = "") -> dict:
    return {"state": state, "overview": overview, "fetched_at": time.time(), "error": error}


def _lg_fetch_one(client: LGClient, settings: LGSettings, file_number: str, session) -> dict:
    status, payload = client.get_json("/api/clm/loan?" + urlencode({"fileNumber": file_number}), session=session)
    if payload is None:
        return _lg_result("not_found") if status == 404 else _lg_result("error", error=f"HTTP {status}")

    items = [item for item in lg_items_of(payload) if isinstance(item, dict)]
    exact = [item for item in items if norm_text(item.get("fileNumber")).casefold() == file_number.casefold()]
    if len(exact) == 1:
        match = exact[0]
    elif len(items) == 1 and not norm_text(items[0].get("fileNumber")):
        match = items[0]  # Some CLM responses omit fileNumber on a filtered lookup.
    elif not items or (not exact and all(norm_text(item.get("fileNumber")) for item in items)):
        return _lg_result("not_found")
    else:
        return _lg_result("error", error="Ambiguous loan match")
    loan_id = match.get("id")
    if not loan_id:
        return _lg_result("not_found")

    template_path = f"/api/clm/pipelineReportTemplates/{quote(str(settings.template_id), safe='')}/loans/{quote(str(loan_id), safe='')}"
    template_status, template = client.get_json(template_path, session=session)
    summary_status, summary = client.get_json(f"/api/clm/loan/{quote(str(loan_id), safe='')}", session=session)
    if template is None and summary is None:
        return _lg_result("error", error=f"Loan details HTTP {max(template_status, summary_status)}")

    overview = {"lg_loan_id": loan_id, "file_number": file_number, "template_ok": template is not None}
    overview.update(lg_parse_template_loan(template) if template is not None else {})
    data = summary.get("data", {}) if isinstance(summary, dict) else {}
    data = data if isinstance(data, dict) else {}
    overview.update(
        balance=data.get("loanBalance"),
        to_finish=data.get("balanceToFinish"),
        last_approved_draw=_lg_date(data.get("lastApprovedDrawEffectiveDate")),
    )
    result = _lg_result("ok", overview)
    result["partial_errors"] = [label + f" (HTTP {code})" for label, code, payload in
                                (("Construction details unavailable", template_status, template),
                                 ("Loan balances unavailable", summary_status, summary)) if payload is None]
    return result


def lg_cached_result(file_number: Optional[str], now: Optional[float] = None, settings: Optional[LGSettings] = None) -> Optional[dict]:
    """A still-fresh cached result for one file number, or None."""
    if not file_number:
        return None
    store = _lg_account_store(settings)
    with store["lock"]:
        failure = store.get("read_failures", {}).get(file_number)
        hit = store["loans"].get(file_number)
    if failure and (now or time.time()) - failure["checked_at"] < LG_ERROR_RETRY_SECONDS:
        return failure
    if not hit:
        return None
    ttl = LG_ERROR_RETRY_SECONDS if hit.get("partial_errors") else LG_CACHE_SECONDS if hit["state"] == "ok" else LG_NOT_FOUND_SECONDS
    return hit if (now or time.time()) - hit["fetched_at"] < ttl else None


def lg_clear_login_backoff(settings=None) -> None:
    store = _lg_account_store(settings)
    with store["lock"]:
        store["login_failure"] = None


def lg_fetch_loans(settings: LGSettings, file_numbers: Iterable[Optional[str]], *, force: bool = False, session_factory=None) -> Dict[str, dict]:
    """{file number: {"state": ok|not_found|error, "overview", "fetched_at", "error"}}.
    Fresh cache entries are reused unless force; the rest are fetched in parallel."""
    store = _lg_account_store(settings)
    wanted = list(dict.fromkeys(fn for fn in file_numbers if fn))
    results: Dict[str, dict] = {}
    need: List[str] = []
    for file_number in wanted:
        hit = None if force else lg_cached_result(file_number, settings=settings)
        if hit is not None:
            results[file_number] = hit
        else:
            need.append(file_number)
    if not need:
        return results

    with store["lock"]:
        failure = store.get("login_failure")
    if failure and not force and time.time() - failure[0] < LG_LOGIN_BACKOFF_SECONDS:
        for file_number in need:
            results[file_number] = {"state": "error", "overview": None, "fetched_at": failure[0], "error": failure[1]}
        return results

    _lg_install_truststore()
    make_session = session_factory or _lg_new_session
    client = LGClient(settings, session=make_session())
    try:
        client.token()  # pre-warm so the workers don't race to log in
    except Exception as exc:
        error = lg_error_text(exc)
        with store["lock"]:
            store["login_failure"] = (time.time(), error)
            for file_number in need:
                previous = store["loans"].get(file_number)
                result = ({**previous, "stale": True, "refresh_error": error}
                          if previous and previous["state"] == "ok" else _lg_result("error", error=error))
                results[file_number] = {**result, "checked_at": time.time()}
                store.setdefault("read_failures", {})[file_number] = results[file_number]
        _lg_close_sessions([client._session])
        return results

    with store["lock"]:
        store["login_failure"] = None

    # Bound the whole batch so an unavailable API cannot block the workbook view.
    client._read_deadline = time.monotonic() + LG_BATCH_READ_SECONDS
    local = threading.local()
    worker_sessions = []

    def work(file_number: str) -> Tuple[str, dict]:
        session = getattr(local, "session", None)
        if session is None:
            session = local.session = make_session()
            worker_sessions.append(session)
        try:
            return file_number, _lg_fetch_one(client, settings, file_number, session)
        except Exception as exc:
            return file_number, _lg_result("error", error=lg_error_text(exc))

    try:
        with ThreadPoolExecutor(max_workers=max(1, min(LG_MAX_WORKERS, len(need)))) as pool:
            for file_number, result in pool.map(work, need):
                results[file_number] = result
    finally:
        _lg_close_sessions([client._session, *worker_sessions])

    with store["lock"]:
        failures = store.setdefault("read_failures", {})
        for file_number in need:
            result = results[file_number]
            if result["state"] in ("ok", "not_found"):
                store["loans"][file_number] = result
                failures.pop(file_number, None)
            else:
                previous = store["loans"].get(file_number)
                if previous and previous["state"] == "ok":
                    result = {**previous, "stale": True, "refresh_error": result["error"]}
                result = {**result, "checked_at": time.time()}
                failures[file_number] = result
                results[file_number] = result
        if len(failures) > LG_STORE_MAX_LOANS:
            for key in sorted(failures, key=lambda k: failures[k]["checked_at"])[:len(failures) - LG_STORE_MAX_LOANS]:
                failures.pop(key, None)
        if len(store["loans"]) > LG_STORE_MAX_LOANS:
            oldest = sorted(store["loans"], key=lambda key: store["loans"][key]["fetched_at"])
            for key in oldest[: len(store["loans"]) - LG_STORE_MAX_LOANS]:
                store["loans"].pop(key, None)
    return results


def lg_loan_column_options(deck: pd.DataFrame) -> List[str]:
    """The default Deal Number plus every workbook column that has values on Bridge rows."""
    bridge = deck[deck["sheet"] == "Bridge"] if "sheet" in deck.columns else deck.iloc[0:0]
    internal = {
        "sheet", "sheet_order", "original_order", "_excel_row", "deal_key", "deal_name", "borrower",
        "account_display", "servicer", "portfolio", "segment", "financing", "loan_buyer", "owner", "status",
        "commentary", "saved_at", "upb", "maturity_date", "next_payment_date", "npl_raw", "days_past_due",
        "funded_amount", "commitment", "commitment_source", "remaining_commitment", "loan_amount",
        "days_to_maturity", "days_to_next_payment", LG_DEFAULT_LOAN_COLUMN,
    }
    options = [LG_DEFAULT_LOAN_COLUMN]
    for col in bridge.columns:
        name = str(col)
        if name in internal or name.startswith(("loan_mod_", "status_")):
            continue
        if bridge[col].map(norm_text).ne("").any():
            options.append(name)
    return options


def lg_deal_file_numbers(deck: pd.DataFrame, loan_column: str, prefix: str) -> Dict[str, Optional[str]]:
    bridge = deck[deck["sheet"] == "Bridge"] if "sheet" in deck.columns else deck.iloc[0:0]
    source = loan_column if loan_column in bridge.columns else LG_DEFAULT_LOAN_COLUMN
    return {
        str(deal_key): lg_file_number(value, prefix)
        for deal_key, value in zip(bridge["deal_key"], safe_series(bridge, source, ""))
    }


def lg_refresh_cb(file_number: Optional[str]) -> None:
    settings = lg_settings()
    if settings is None or not file_number:
        return
    lg_clear_login_backoff(settings)
    lg_fetch_loans(settings, [file_number], force=True)


def lg_refresh_all_cb() -> None:
    lg_clear_login_backoff()
    st.session_state["_lg_force_all"] = True


def prepare_landgorilla_context(deck: pd.DataFrame) -> Optional[LGContext]:
    """Read available snapshots only. Navigation must never perform API requests."""
    settings = lg_settings()
    if settings is None or deck.empty:
        return None

    options = lg_loan_column_options(deck)
    if st.session_state.get("lg_loan_column") not in options:
        st.session_state["lg_loan_column"] = LG_DEFAULT_LOAN_COLUMN
    loan_column = st.session_state["lg_loan_column"]

    deal_files = lg_deal_file_numbers(deck, loan_column, settings.file_prefix)
    store = _lg_account_store(settings)
    results = {}
    now = time.time()
    with store["lock"]:
        for fn in set(deal_files.values()) - {None}:
            hit = store.get("read_failures", {}).get(fn) or store["loans"].get(fn)
            failure = store.get("login_failure")
            if hit is None and failure and now - failure[0] < LG_LOGIN_BACKOFF_SECONDS:
                hit = {"state": "error", "overview": None, "fetched_at": failure[0], "error": failure[1]}
            if hit is not None:
                # Retain an expired snapshot for reference, with its original time.
                # Only the selected loan's explicit Load/Refresh button uses the API.
                results[fn] = {**hit, "cache_expired": now - hit["fetched_at"] >= LG_CACHE_SECONDS}
    return LGContext(settings=settings, loan_column=loan_column, deal_files=deal_files, results=results)


def lg_result_for_row(row: Mapping[str, object], lg_ctx: Optional[LGContext]) -> Tuple[Optional[str], Optional[dict]]:
    if lg_ctx is None:
        return None, None
    file_number = lg_ctx.deal_files.get(str(row.get("deal_key")))
    return file_number, (lg_ctx.results.get(file_number) if file_number else None)


def lg_fetched_text(result: Optional[dict]) -> str:
    if not result:
        return ""
    stamp = pd.Timestamp(result["fetched_at"], unit="s", tz="UTC").tz_convert(LG_TIMEZONE)
    return stamp.strftime("%I:%M %p ET").lstrip("0")


def md_escape(text: object) -> str:
    """Escape text for st.markdown / alerts so '$' is not read as LaTeX and markup stays literal."""
    return re.sub(r"([\\`*_{}\[\]<>()#+!|~$])", r"\\\1", str(text))


def lg_presenter_prompts(row: pd.Series, lg_ctx: Optional[LGContext]) -> List[str]:
    _, result = lg_result_for_row(row, lg_ctx)
    if not result or result["state"] != "ok":
        return []
    overview = result["overview"]
    checks = lg_checks(row, overview, lg_derived(overview, lg_today()), lg_today())
    return [f"Land Gorilla: {message}" for level, message in checks if level in ("alert", "warn")][:2]


def render_lg_checks(checks: List[Tuple[str, str]]) -> None:
    """One compact box per severity, so a flagged loan doesn't push the slide off screen."""
    for level, render, icon in (
        ("alert", st.error, ":material/error:"),
        ("warn", st.warning, ":material/warning:"),
        ("info", st.info, ":material/info:"),
    ):
        messages = [message for check_level, message in checks if check_level == level]
        if not messages:
            continue
        if len(messages) == 1:
            render(md_escape(messages[0]), icon=icon)
        else:
            render("\n".join(f"- {md_escape(message)}" for message in messages), icon=icon)


def render_landgorilla_panel(current_row: pd.Series, lg_ctx: Optional[LGContext]) -> None:
    """Optional construction snapshot, loaded only by an explicit user action."""
    if lg_ctx is None or norm_text(current_row.get("sheet")) != "Bridge":
        return
    file_number, result = lg_result_for_row(current_row, lg_ctx)
    deal_key = str(current_row.get("deal_key"))
    with st.expander("Construction · Land Gorilla", expanded=False):
        details = [f"File {file_number}" if file_number else "No Land Gorilla loan number"]
        if result:
            stamp = pd.Timestamp(result["fetched_at"], unit="s", tz="UTC").tz_convert(LG_TIMEZONE)
            details.append(f"Retrieved {stamp:%m/%d/%Y %I:%M %p} ET")
        st.caption(md_escape(" · ".join(details)))
        st.button(
            "Refresh construction data" if result else "Load construction data",
            key=f"lg_refresh::{deal_key}", icon=":material/refresh:",
            on_click=lg_refresh_cb, args=(file_number,), disabled=not file_number,
            help="Request only this loan. Moving between slides never calls Land Gorilla.",
        )
        if not file_number:
            st.caption("Select the loan-number column under Land Gorilla in the sidebar.")
            return
        if result is None:
            st.caption("Load this loan when needed. Workbook figures are shown above.")
            return
        if result["state"] == "error":
            st.caption(f"Construction data could not be loaded ({md_escape(result.get('error') or 'no response')}). Workbook figures remain available.")
            return
        if result["state"] == "not_found":
            st.caption("No matching Land Gorilla file was returned. Workbook figures remain available.")
            return
        if result.get("stale"):
            st.caption(f"Refresh failed ({md_escape(result.get('refresh_error') or 'no response')}). These are the last retrieved values; check the retrieval time above.")
        elif result.get("cache_expired"):
            st.caption("Saved snapshot; refresh to check for newer values.")
        for message in result.get("partial_errors", []):
            st.caption(md_escape(message) + ". Available fields are shown below.")
        overview = result["overview"]
        derived = lg_derived(overview, lg_today())
        project, duration = derived["project_pct"], derived["duration_pct"]
        left, right = st.columns(2)
        if project is not None:
            left.progress(project / 100, text=f"Construction {project:.0f}% complete")
        if duration is not None:
            right.progress(duration / 100, text=f"Schedule {duration:.0f}% elapsed")
        cols = st.columns(4)
        cols[0].metric("LG balance", fmt_money_exact(derived["balance"]))
        cols[1].metric("Construction holdback remaining", fmt_money_exact(derived["to_finish"]),
                       help="Land Gorilla balanceToFinish; this is not the full unfunded loan commitment.")
        cols[2].metric("Last approved draw", fmt_date(overview.get("last_approved_draw")))
        cols[3].metric("LG due date", fmt_date(overview.get("due_date")))
        if overview.get("status"):
            st.caption("LG status: " + md_escape(overview["status"]))
        labels = [label for label in overview.get("risk") or [] if label]
        if labels:
            st.caption("LG labels / instructions: " + md_escape(" · ".join(labels)))
        # Source differences are observations, not automatic credit or servicing alerts.
        comparisons = []
        balance, upb = derived["balance"], lg_number(current_row.get("upb"))
        if balance is not None and upb is not None and abs(balance - upb) > max(LG_BALANCE_TOLERANCE, abs(upb) * LG_BALANCE_TOLERANCE_PCT):
            comparisons.append(f"LG balance {fmt_money_exact(balance)}; workbook UPB {fmt_money_exact(upb)}")
        due, maturity = overview.get("due_date"), coerce_datetime_value(current_row.get("maturity_date"))
        if isinstance(due, dt.date) and not pd.isna(maturity) and abs((pd.Timestamp(due) - pd.Timestamp(maturity).normalize()).days) > LG_DUE_DATE_TOLERANCE_DAYS:
            comparisons.append(f"LG due date {fmt_date(due)}; workbook maturity {fmt_date(maturity)}")
        for comparison in comparisons:
            st.caption(md_escape(comparison) + ". Sources may have different effective dates.")


def build_landgorilla_table(deck: pd.DataFrame, lg_ctx: LGContext, today: dt.date) -> pd.DataFrame:
    rows = []
    for _, row in deck[deck["sheet"] == "Bridge"].iterrows():
        _, result = lg_result_for_row(row, lg_ctx)
        if not result or result["state"] != "ok":
            continue
        overview = result["overview"]
        derived = lg_derived(overview, today)
        checks = lg_checks(row, overview, derived, today)
        rows.append(
            {
                "Deal Name": display_text(row.get("deal_name")),
                "Deal #": display_text(row.get("deal_number")),
                "Built": derived["project_pct"],
                "Schedule elapsed": derived["duration_pct"],
                "Behind (pts)": derived["gap_pts"],
                "Last draw": overview.get("last_draw"),
                "Days since draw": derived["days_since_draw"],
                "Left to fund": derived["to_finish"],
                "LG balance": derived["balance"],
                "Workbook UPB": lg_number(row.get("upb")),
                "LG due date": overview.get("due_date"),
                "LG status": overview.get("status") or "",
                "Retrieved at": pd.Timestamp(result["fetched_at"], unit="s", tz="UTC").tz_convert(LG_TIMEZONE).strftime("%m/%d/%Y %I:%M %p ET"),
                "Source freshness": "Last known" if result.get("stale") or result.get("cache_expired") else "Partial response" if result.get("partial_errors") else "Saved snapshot",
                "LG labels / instructions": list(overview.get("risk") or []),
                "Checks": sum(1 for level, _ in checks if level in ("alert", "warn")),
            }
        )
    frame = pd.DataFrame(rows)
    if not frame.empty:
        frame = frame.sort_values("Behind (pts)", ascending=False, na_position="last", kind="stable")
    return frame


def render_landgorilla_overview(deck: pd.DataFrame, lg_ctx: Optional[LGContext]) -> None:
    if lg_ctx is None or deck.empty or not (deck["sheet"] == "Bridge").any():
        return

    table = build_landgorilla_table(deck, lg_ctx, lg_today())
    bridge_count = int((deck["sheet"] == "Bridge").sum())
    if table.empty:
        return
    st.subheader("Construction snapshots · Land Gorilla")
    st.caption(f"Saved data for {len(table)} of {bridge_count} Bridge loans. Load or refresh a loan from Presentation when needed.")
    st.dataframe(
        table.drop(columns=["Behind (pts)", "Checks", "Days since draw"], errors="ignore"),
        hide_index=True,
        width="stretch",
        column_config={
            "Built": st.column_config.ProgressColumn("Built", min_value=0, max_value=100, format="%.0f%%"),
            "Schedule elapsed": st.column_config.ProgressColumn("Schedule elapsed", min_value=0, max_value=100, format="%.0f%%"),
            "Behind (pts)": st.column_config.NumberColumn("Behind (pts)", format="%.0f", help="Schedule elapsed minus construction complete"),
            "Last draw": st.column_config.DateColumn("Last draw", format="MM/DD/YYYY"),
            "Days since draw": st.column_config.NumberColumn("Days since draw", format="%d"),
            "Left to fund": st.column_config.NumberColumn("Left to fund", format="dollar"),
            "LG balance": st.column_config.NumberColumn("LG balance", format="dollar"),
            "Workbook UPB": st.column_config.NumberColumn("Workbook UPB", format="dollar"),
            "LG due date": st.column_config.DateColumn("LG due date", format="MM/DD/YYYY"),
            "LG labels / instructions": st.column_config.ListColumn("LG labels / instructions"),
            "Checks": st.column_config.NumberColumn("Checks", help="Mismatches or warnings vs. the workbook", format="%d"),
        },
    )


def render_landgorilla_sidebar(lg_ctx: Optional[LGContext], deck: pd.DataFrame) -> None:
    with st.sidebar.expander("Land Gorilla"):
        if lg_ctx is None:
            st.caption("Not configured. Optional construction data uses the app's Land Gorilla secrets.")
            return
        st.caption("Construction data loads on request for the selected Bridge loan. Navigation uses saved snapshots only.")
        st.selectbox(
            "Loan # column", options=lg_loan_column_options(deck), key="lg_loan_column",
            format_func=lambda col: "Deal Number (default)" if col == LG_DEFAULT_LOAN_COLUMN else col,
            help=f"Land Gorilla file numbers are '{lg_ctx.settings.file_prefix}' + this number.",
        )
