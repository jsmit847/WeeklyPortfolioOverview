import datetime as dt

import pandas as pd
import pytest

import portfolio_cockpit as pc


def sample_row(**overrides):
    row = dict(deal_key="Term::1", deal_name="Synthetic loan", deal_number="1", sheet="Term", owner="Alice",
               status="Current", portfolio="Sold Term", upb=1_000_000, commitment=2_000_000, days_past_due=0,
               days_to_next_payment=-522, days_to_maturity=200, maturity_date=dt.date(2027, 5, 1),
               loan_mod_count=0, loan_mod_active_count=0, status_needs_update_prompt=False)
    row.update(overrides)
    return row


def test_sold_loan_does_not_invent_payment_overdue():
    row = sample_row()
    assert "STALE_PAYMENT_DATE" in {x["code"] for x in pc.consistency_checks(row)}
    flags = pc.compute_flags(pd.DataFrame([row]))
    assert "Payment overdue" not in flags.Rule.tolist()
    assert "Date not current?" in " ".join(flags.Detail)


def test_current_dpd_zero_with_old_payment_is_data_check():
    row = sample_row(portfolio="Term")
    assert "STALE_PAYMENT_DATE" in {x["code"] for x in pc.consistency_checks(row)}
    assert "Payment overdue" not in pc.compute_flags(pd.DataFrame([row])).Rule.tolist()


def test_true_overdue_requires_supporting_dpd():
    row = sample_row(portfolio="Term", days_past_due=35)
    assert {"Payment overdue", "30+ DPD"} <= set(pc.compute_flags(pd.DataFrame([row])).Rule)


@pytest.mark.parametrize("days,bucket", [(-1, "Past maturity"), (0, "0–30 days"), (30, "0–30 days"),
    (31, "31–60 days"), (61, "61–90 days"), (91, "91–180 days"), (181, "181–365 days"),
    (366, ">1 year"), (None, "Unknown"), (float("nan"), "Unknown")])
def test_maturity_boundaries(days, bucket):
    assert pc.maturity_bucket(days) == bucket


def test_wall_conserves_balance():
    deck = pd.DataFrame([sample_row(), sample_row(deal_key="Bridge::2", sheet="Bridge", days_to_maturity=-1)])
    wall = pc.maturity_wall(deck)
    assert wall.UPB.sum() == 2_000_000
    assert wall.Bucket.tolist()[0] == "Past maturity"


def test_attention_order_is_stable_and_does_not_change_keys():
    deck = pd.DataFrame([sample_row(deal_key="a"), sample_row(deal_key="b", days_to_maturity=-2), sample_row(deal_key="c")])
    assert pc.order_deck(deck, "Attention first").deal_key.tolist() == ["b", "a", "c"]
    assert deck.deal_key.tolist() == ["a", "b", "c"]


def test_owner_order_unassigned_last():
    deck = pd.DataFrame([sample_row(deal_key="a", owner=""), sample_row(deal_key="b", owner="Zoe"), sample_row(deal_key="c", owner="Alice")])
    assert pc.order_deck(deck, "By owner").deal_key.tolist() == ["c", "b", "a"]


def test_safe_csv_preserves_numeric_values_neutralizes_formulas():
    data = pc.safe_csv(pd.DataFrame({"Text": ["=1+1", "  @bad"], "Amount": [-12, 1]})).decode("utf-8-sig")
    assert "'=1+1" in data and "'  @bad" in data and "-12" in data


def test_commentary_rule_dismisses_status_only_reminder_when_update_changed():
    deck = pd.DataFrame([sample_row(status_needs_update_prompt=True, commentary="New context")])
    history = pd.DataFrame([{"deal_key": "Term::1", "meeting_date": date, "commentary": "Older context"}
                            for date in ("2026-10-01", "2026-10-08")])
    assert not pc.apply_reminder_rule(deck, history, dt.date(2026, 10, 15), "Status + commentary").iloc[0].status_needs_update_prompt
    assert pc.apply_reminder_rule(deck, history, dt.date(2026, 10, 15), "Status only").iloc[0].status_needs_update_prompt


def test_export_verifier_checks_edited_cells(app_module, sample_bytes):
    parsed = app_module.parse_workbook_bytes(sample_bytes)
    base = app_module.build_base_rows(parsed)
    key = next(iter(base))
    item = {**base[key], "status": "New synthetic status", "owner": "", "commentary": "=literal\nsecond line"}
    exported = app_module.update_workbook_bytes(sample_bytes, {key: item})
    result = pc.verify_workbook(sample_bytes, exported, {key: item}, app_module)
    assert result["edits"] == 1
    with pytest.raises(ValueError, match="does not match"):
        pc.verify_workbook(sample_bytes, exported, {key: {**item, "status": "Different"}}, app_module)
