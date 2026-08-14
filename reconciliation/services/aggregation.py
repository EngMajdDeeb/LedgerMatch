"""
Deterministic (non-AI) arithmetic helpers.

Summing numbers the vision step already extracted, or numbers already
present in the ledger, is plain arithmetic - not the kind of judgement call
("is this the same item, is this difference meaningful") that the brief asks
to leave to the AI comparison step. Keeping it here keeps totals exact and
keeps the AI call focused on categorization/matching/narrative.
"""
from __future__ import annotations

from ..constants import BALANCE_CATEGORY_KEYS, FLOW_CATEGORY_KEYS


def aggregate_image_totals(extraction: dict) -> dict:
    """Sum per-shift flow categories and resolve the day-level balance fields.

    Opening/closing balances, the pharmacy's transcription of the El-Bayan
    balance, and its self-computed discrepancy are day-level snapshots that
    the vision step is asked to capture twice: once in the top-level
    `day_totals` (wherever they appear on the page) and once inside each
    shift's `values` (in case they happen to sit under a shift's column).
    `day_totals` is preferred; if the model left it null, fall back to the
    last non-null value seen across the shifts.
    """
    totals = {key: None for key in FLOW_CATEGORY_KEYS + BALANCE_CATEGORY_KEYS}
    shifts = extraction.get('shifts') or []

    for key in FLOW_CATEGORY_KEYS:
        values = [
            shift.get('values', {}).get(key)
            for shift in shifts
            if shift.get('values', {}).get(key) is not None
        ]
        if values:
            totals[key] = round(sum(values), 2)

    day_totals = extraction.get('day_totals') or {}
    for key in BALANCE_CATEGORY_KEYS:
        value = day_totals.get(key)
        if value is None:
            for shift in reversed(shifts):
                shift_value = shift.get('values', {}).get(key)
                if shift_value is not None:
                    value = shift_value
                    break
        totals[key] = value

    return totals
