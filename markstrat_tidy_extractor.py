import pandas as pd
import numpy as np
import re

# ---------- helpers ----------
def _find_anchor(df: pd.DataFrame, text: str):
    coords = np.where(df.astype(str).applymap(lambda x: isinstance(x, str) and text in x).values)
    if len(coords[0]) == 0:
        return None
    return int(coords[0][0]), int(coords[1][0])

def _detect_period_anywhere(df: pd.DataFrame) -> int | None:
    period_found = None
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = df.iat[r, c]
            if isinstance(val, str) and val.startswith("Period "):
                m = re.search(r"Period\s*(\d+)", val)
                if m:
                    period_found = int(m.group(1))
    return period_found

def _is_number(x):
    return isinstance(x, (int, float, np.integer, np.floating)) and pd.notna(x)

# ---------- extractors ----------
def extract_product_contribution_to_tidy(xlsx_path: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path, sheet_name="Firm", header=None)
    sheet = "Firm"
    unit = None

    anchor = _find_anchor(df, "Product Contribution")
    if not anchor:
        raise ValueError("Could not find 'Product Contribution' section in Firm sheet.")
    anchor_r, anchor_c = anchor

    window = df.iloc[anchor_r:anchor_r+30, max(0, anchor_c-1):anchor_c+20].copy()
    base_r = anchor_r
    base_c = max(0, anchor_c-1)

    known = [
        "Revenues",
        "Cost of goods sold",
        "Inventory holding costs",
        "Inventory selling costs",
        "Contribution before marketing",
        "Advertising media",
        "Advertising research",
        "Commercial team costs",
        "Contribution after marketing",
    ]

    metric_rows = []
    for r in range(1, window.shape[0]):
        cell = window.iat[r, 3] if window.shape[1] > 3 else None
        if isinstance(cell, str) and cell.strip():
            txt = cell.strip()
            if any(k.lower() in txt.lower() for k in known):
                metric_rows.append((r, 3, txt))

    for r in range(window.shape[0]):
        for c in range(window.shape[1]):
            val = window.iat[r, c]
            if isinstance(val, str) and "thousands of dollars" in val.lower():
                unit = "thousands of dollars"

    product_row_guess = None
    for r in range(len(window)-1, -1, -1):
        row_vals = window.iloc[r].tolist()
        caps = [v for v in row_vals if isinstance(v, str) and re.match(r"^[A-Z0-9]{3,}$", v.strip())]
        if len(caps) >= 2:
            product_row_guess = r
            break

    col_to_product = {}
    if product_row_guess is not None:
        for c in range(window.shape[1]):
            v = window.iat[product_row_guess, c]
            if isinstance(v, str) and re.match(r"^[A-Z0-9]{3,}$", v.strip()):
                col_to_product[c] = v.strip()

    period = _detect_period_anywhere(df)

    records = []
    for (mr, mc, metric_text) in metric_rows:
        for c in range(window.shape[1]):
            val = window.iat[mr, c]
            if _is_number(val):
                product = col_to_product.get(c)
                if product:
                    records.append({
                        "section": "Product Contribution",
                        "period": period,
                        "product": product,
                        "metric": metric_text,
                        "value": float(val),
                        "unit": unit,
                        "sheet": sheet,
                        "row": base_r + mr,
                        "col": base_c + c,
                    })

    return pd.DataFrame.from_records(records)

def extract_company_pl_to_tidy(xlsx_path: str) -> pd.DataFrame:
    sheet = "Firm"
    unit = "thousands of dollars"
    df = pd.read_excel(xlsx_path, sheet_name=sheet, header=None)

    anchor = _find_anchor(df, "Company Profit & Loss Statement")
    if not anchor:
        raise Value
