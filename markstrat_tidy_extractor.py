import pandas as pd
import numpy as np
import re

# ============ helpers ============
def _find_anchor(df, text):
    coords = np.where(df.astype(str).applymap(lambda x: isinstance(x, str) and text in x).values)
    if len(coords[0]) == 0:
        return None
    return int(coords[0][0]), int(coords[1][0])

def _detect_period_anywhere(df):
    period_found = None
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = df.iat[r, c]
            if isinstance(val, str) and val.startswith("Period "):
                m = re.search(r"Period\\s*(\\d+)", val)
                if m:
                    period_found = int(m.group(1))
    return period_found

def _is_number(x):
    return isinstance(x, (int, float, np.integer, np.floating)) and pd.notna(x)

# ============ extractors ============
def extract_product_contribution_to_tidy(xlsx_path):
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

def extract_company_pl_to_tidy(xlsx_path):
    sheet = "Firm"
    unit = "thousands of dollars"
    df = pd.read_excel(xlsx_path, sheet_name=sheet, header=None)

    anchor = _find_anchor(df, "Company Profit & Loss Statement")
    if not anchor:
        raise ValueError("Company Profit & Loss Statement not found.")
    ar, ac = anchor

    win = df.iloc[ar:ar+60, max(0, ac-1):ac+40].copy()
    base_c = max(0, ac-1)

    header_row = None
    period_cols = {}
    for r in range(win.shape[0]):
        tmp = {}
        for c in range(win.shape[1]):
            v = win.iat[r, c]
            if isinstance(v, str):
                s = v.strip()
                m = re.match(r"^Period\\s*(\\d+)$", s)
                if m:
                    tmp[c] = int(m.group(1))
                elif s.lower() == "cumulative":
                    tmp[c] = "Cumulative"
        if len(tmp) >= 3:
            header_row = r
            period_cols = tmp
            break
    if header_row is None:
        raise ValueError("Could not locate the Period header row in P&L.")

    records = []
    for r in range(header_row + 1, win.shape[0]):
        row_texts = [win.iat[r, c] for c in range(win.shape[1])]
        joined = " ".join([str(x) for x in row_texts if isinstance(x, str)]).lower()
        if "market contribution" in joined or "all numbers" in joined:
            break
        if all(pd.isna(x) for x in row_texts):
            continue

        metric = None
        for c in range(win.shape[1]):
            v = win.iat[r, c]
            if isinstance(v, str) and v.strip():
                metric = v.strip()
                break
        if not metric:
            continue

        for c, per_label in period_cols.items():
            if c >= win.shape[1]:
                continue
            val = win.iat[r, c]
            if _is_number(val):
                records.append({
                    "section": "Company Profit & Loss Statement",
                    "metric": metric,
                    "period": per_label,
                    "value": float(val),
                    "unit": unit,
                    "sheet": sheet,
                    "row": ar + r,
                    "col": base_c + c,
                })

    return pd.DataFrame.from_records(records)
