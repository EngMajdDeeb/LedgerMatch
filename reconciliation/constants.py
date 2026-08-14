"""
Shared vocabulary for the cash-box (صندوق) reconciliation categories.

Every pharmacy sends an end-of-day cash-box summary (opening balance, sales,
returns, purchases, supplier payments, expenses, closing balance...). The
El-Bayan export is a raw transaction ledger for the same day that has to be
categorized into the same buckets before the two sides can be compared.

This module is the single source of truth for those category keys/labels so
the vision prompt, the comparison prompt, the Excel audit export and the
templates all agree on the same shape.
"""

# Categories that accumulate across the day (sums of transactions/entries).
FLOW_CATEGORIES = [
    ("sales", "مبيعات"),
    ("sales_returns", "مرتجع مبيعات"),
    ("cash_purchases", "مشتريات / نقدا"),
    ("credit_purchases", "مشتريات / آجل"),
    ("purchase_returns", "مرتجع مشتريات"),
    ("supplier_payments", "دفعات موردين"),
    ("company_funding", "تمويل شركة"),
    ("expenses_1", "مصاريف 1"),
    ("expenses_2", "مصاريف 2"),
    ("expenses_3", "مصاريف 3"),
]

# Categories that represent a point-in-time balance/snapshot rather than a
# sum of movements during the day.
BALANCE_CATEGORIES = [
    ("opening_balance", "صندوق بداية"),
    ("closing_balance", "صندوق نهاية"),
    ("bayan_balance", "صندوق البيان"),
    ("discrepancy", "فروقات الصندوق"),
]

ALL_CATEGORIES = FLOW_CATEGORIES + BALANCE_CATEGORIES

CATEGORY_LABELS_AR = dict(ALL_CATEGORIES)

FLOW_CATEGORY_KEYS = [key for key, _ in FLOW_CATEGORIES]
BALANCE_CATEGORY_KEYS = [key for key, _ in BALANCE_CATEGORIES]
ALL_CATEGORY_KEYS = [key for key, _ in ALL_CATEGORIES]
