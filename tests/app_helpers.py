"""Helpers for AppTest-based tests (import from tests: `from app_helpers import ...`)."""
from __future__ import annotations

import html
import re
import sys
from typing import Dict, List, NamedTuple

from streamlit.testing.v1 import AppTest

COUNTERS_MODULE = "wpo_test_counters"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def call_counts() -> Dict[str, int]:
    return dict(getattr(sys.modules.get(COUNTERS_MODULE), "calls", {}) or {})


def goto_view(at: AppTest, view: str) -> AppTest:
    at.sidebar.radio(key="view_mode").set_value(view).run()
    return at


def kpi_html(at: AppTest) -> str:
    """The rendered KPI strip markup (not the <style> block that also names its class)."""
    for element in at.markdown:
        if "<div class='rt-kpi-grid'>" in str(element.value):
            return str(element.value)
    return ""


class KpiTile(NamedTuple):
    label: str
    value: str
    helper: str
    signal: str
    tooltip: str


_KPI_TILE_RE = re.compile(
    r"<div class='rt-kpi'(?: title='(?P<tooltip>[^']*)')?>"
    r"<div class='rt-kpi-label'>(?P<label>.*?)</div>"
    r"<div class='rt-kpi-value'>(?P<value>.*?)</div>"
    r"(?:<div class='rt-kpi-helper (?P<signal>rt-signal-\w+)'>(?P<helper>.*?)</div>)?"
    r"</div>"
)


def kpi_tiles(at: AppTest) -> Dict[str, KpiTile]:
    """{label: KpiTile} parsed from the KPI strip (text unescaped)."""
    tiles: Dict[str, KpiTile] = {}
    for match in _KPI_TILE_RE.finditer(kpi_html(at)):
        tile = KpiTile(
            html.unescape(match.group("label")),
            html.unescape(match.group("value")),
            html.unescape(match.group("helper") or ""),
            match.group("signal") or "",
            html.unescape(match.group("tooltip") or ""),
        )
        tiles[tile.label] = tile
    return tiles


def all_markdown(at: AppTest) -> str:
    return "\n".join(str(element.value) for element in at.markdown)


def metric_by_label(at: AppTest, label: str):
    for metric in at.metric:
        if metric.label == label:
            return metric
    raise AssertionError(f"no metric {label!r}; have {[m.label for m in at.metric]}")


def button_by_label(at: AppTest, label: str):
    for button in at.button:
        if button.label == label:
            return button
    raise AssertionError(f"no button {label!r}; have {[b.label for b in at.button]}")


def progress_text(at: AppTest) -> str:
    return at.get("progress")[0].proto.text if at.get("progress") else ""


def jump_to(at: AppTest, deal_number: str) -> AppTest:
    """Use the Presentation 'Jump to deal' selectbox to open a deal by number."""
    box = next(sb for sb in at.selectbox if sb.label == "Jump to deal")
    option = next(opt for opt in box.options if opt.endswith(f"| {deal_number}"))
    box.set_value(option).run()
    return at


def warnings_text(at: AppTest, *, sidebar: bool = True) -> List[str]:
    source = at.sidebar.warning if sidebar else at.warning
    return [str(w.value) for w in source]


def upload_workbook(at: AppTest, data: bytes, name: str = "other_week.xlsx") -> AppTest:
    at.sidebar.file_uploader(key="uploaded_workbook").upload(name, data, XLSX_MIME).run()
    return at
