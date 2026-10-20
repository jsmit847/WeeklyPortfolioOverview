"""Shared test fixtures.

- Streamlit caches are cleared around every test (autouse).
- app_module: portfolio_meeting_app imported with importlib; main() does not run.
- sample_bytes / scaled_bytes: synthetic workbooks pinned to as-of 2026-10-15.
- make_at(...): writes a workbook as portfolio_overview.xlsx into a temp working folder,
  points WPO_DEFAULT_WORKBOOK at it and returns AppTest.from_file(default_timeout=120).
- patched_at(...): an AppTest.from_function harness that imports the app as its own
  module, wraps named functions to count calls (or raise), then calls main().
- goto(at, view): switch the sidebar View radio and run.
Only synthetic data is used anywhere in the suite.

Set WPO_APP_PATH to run the suite against another copy of the app, e.g. an older commit
extracted with `git show 10afc49:portfolio_meeting_app.py`, to check that the regression
tests in test_bug_regressions.py fail there.
"""
from __future__ import annotations

import datetime as dt
import importlib.util
import os
import sys
import types
import uuid
from pathlib import Path
from typing import Callable, Dict, Optional, Sequence

import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest

REPO = Path(__file__).resolve().parents[1]
APP_PATH = Path(os.environ.get("WPO_APP_PATH") or REPO / "portfolio_meeting_app.py").resolve()
TESTS_DIR = Path(__file__).resolve().parent
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))

import sample_workbook as sw  # noqa: E402
from app_helpers import goto_view  # noqa: E402

AS_OF = dt.date(2026, 10, 15)
COUNTERS_MODULE = "wpo_test_counters"


def load_app_module(name: str = "wpo_app"):
    spec = importlib.util.spec_from_file_location(name, APP_PATH)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(autouse=True)
def _clear_streamlit_caches():
    st.cache_data.clear()
    st.cache_resource.clear()
    yield
    st.cache_data.clear()
    st.cache_resource.clear()


@pytest.fixture(scope="session")
def app_module():
    return load_app_module()


@pytest.fixture(scope="session")
def sample_bytes() -> bytes:
    return sw.build_sample_workbook_bytes(as_of=AS_OF)


@pytest.fixture(scope="session")
def scaled_bytes() -> bytes:
    return sw.build_scaled_workbook_bytes(n_bridge=120, n_term=30, n_weeks=52, as_of=AS_OF)


def _write_workbook(
    tmp_path: Path,
    monkeypatch,
    workbook_bytes: Optional[bytes],
    file_name: str,
    variants,
    as_of: dt.date = AS_OF,
) -> Path:
    data = workbook_bytes if workbook_bytes is not None else sw.build_sample_workbook_bytes(as_of=as_of, **variants)
    path = tmp_path / file_name
    path.write_bytes(data)
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("WPO_DEFAULT_WORKBOOK", str(path))
    return path


@pytest.fixture
def make_at(tmp_path, monkeypatch) -> Callable[..., AppTest]:
    """make_at(workbook_bytes=None, *, file_name=..., as_of=AS_OF, **variants) -> AppTest.

    Without workbook_bytes the sample is built for `as_of` with the given variant flags.
    The app takes its as-of date from the workbook's title row on the first run, so the
    dates are pinned by the workbook, not by the machine date.
    """

    def _make(
        workbook_bytes: Optional[bytes] = None,
        *,
        file_name: str = "portfolio_overview.xlsx",
        as_of: dt.date = AS_OF,
        **variants,
    ) -> AppTest:
        _write_workbook(tmp_path, monkeypatch, workbook_bytes, file_name, variants, as_of)
        at = AppTest.from_file(str(APP_PATH), default_timeout=120)
        at.session_state['meeting_date'] = AS_OF
        return at

    return _make


def _patched_script(app_path, module_name, count, raise_in):
    # AppTest.from_function runs this function's source as a standalone script, so it
    # carries its own imports and no annotations (Python 3.13 evaluates them eagerly).
    import importlib.util
    import sys
    import types

    counters = sys.modules.setdefault("wpo_test_counters", types.ModuleType("wpo_test_counters"))
    if not hasattr(counters, "calls"):
        counters.calls = {}
    module = sys.modules.get(module_name)
    if module is None:
        spec = importlib.util.spec_from_file_location(module_name, app_path)
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
        sys.modules[module_name] = module
        for name in count:
            original = getattr(module, name)

            def counted(*args, _name=name, _original=original, **kwargs):
                counters.calls[_name] = counters.calls.get(_name, 0) + 1
                return _original(*args, **kwargs)

            setattr(module, name, counted)
        for name in raise_in:

            def boom(*args, _name=name, **kwargs):
                raise RuntimeError(f"injected failure in {_name}")

            setattr(module, name, boom)
    module.main()


@pytest.fixture
def patched_at(tmp_path, monkeypatch):
    created = []
    counters = sys.modules.setdefault(COUNTERS_MODULE, types.ModuleType(COUNTERS_MODULE))
    counters.calls = {}

    def _make(
        workbook_bytes: Optional[bytes] = None,
        *,
        count: Sequence[str] = (),
        raise_in: Sequence[str] = (),
        file_name: str = "portfolio_overview.xlsx",
        **variants,
    ) -> AppTest:
        _write_workbook(tmp_path, monkeypatch, workbook_bytes, file_name, variants)
        module_name = f"wpo_patched_{uuid.uuid4().hex[:8]}"
        created.append(module_name)
        at = AppTest.from_function(
            _patched_script,
            default_timeout=120,
            kwargs={
                "app_path": str(APP_PATH),
                "module_name": module_name,
                "count": tuple(count),
                "raise_in": tuple(raise_in),
            },
        )

        at.session_state['meeting_date'] = AS_OF
        return at

    _make.calls = counters.calls  # type: ignore[attr-defined]
    yield _make
    for name in created:
        sys.modules.pop(name, None)


@pytest.fixture
def goto() -> Callable[[AppTest, str], AppTest]:
    return goto_view
