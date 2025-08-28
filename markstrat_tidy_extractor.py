import pandas as pd
import numpy as np
import re

def extract_product_contribution_to_tidy(xlsx_path: str) -> pd.DataFrame:
    """
    Extracts the Product Contribution block from the 'Firm' sheet
    and returns a tidy table with columns:
    ['period','product','metric','value','unit','sheet','row','col']
    """
    df = pd.read_excel(xlsx_path, sheet_name="Firm", header=None)
    sheet = "Firm"
    unit = None

    # Find "Product Contribution" anchor
    coords = np.where(
        df.astype(str).applymap(lambda x: isinstance(x, str) and "Product Contribution" in x).values
    )
    if len(coords[0]) == 0:
        raise ValueError("Could not find 'Product Contribution' section.")
    anchor_r, anchor_c = int(coords[0][0]), int(coords[1][0])

    # Window around the block
    window = df.iloc[anchor_r:anchor_r+30, max(0, anchor_c-1):anchor_c+12].copy()
    base_r = anchor_r
    base_c = max(0, anchor_c-1)

    # Known metric labels in this block
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

    # Metrics appear in column index 3 of the window in your export
    metric_rows = []
    for r in range(1, window.shape[0]):
        cell = window.iat[r, 3] if window.shape[1] > 3 else None
        if isinstance(cell, str) and cell.strip():
            txt = cell.strip()
            if any(k.lower() in txt.lower() for k in known):
                metric_rows.append((r, 3, txt))

    # Unit (optional)
    for r in range(window.shape[0]):
        for c in range(window.shape[1]):
            val = window.iat[r, c]
            if isinstance(val, str) and "thousands of dollars" in val.lower():
                unit = "thousands of dollars"

    # Product header row: find a row with multiple ALLCAPS tokens (SOCOOL, SOFT, SOLO)
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

    # Period: pick the highest "Period N" string in the sheet
    period_found = None
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = df.iat[r, c]
            if isinstance(val, str) and val.startswith("Period "):
                m = re.search(r"Period\s*(\d+)", val)
                if m:
                    period_found = int(m.group(1))

    # Collect numeric values per metric across product columns
    records = []
    numeric_types = (int, float, np.integer, np.floating)
    for (mr, mc, metric_text) in metric_rows:
        for c in range(window.shape[1]):
            val = window.iat[mr, c]
            if isinstance(val, numeric_types) and pd.notna(val):
                product = col_to_product.get(c)
                if product:
                    records.append({
                        "period": period_found,
                        "product": product,
                        "metric": metric_text,
                        "value": float(val),
                        "unit": unit,
                        "sheet": sheet,
                        "row": base_r + mr,
                        "col": base_c + c
                    })

    return pd.DataFrame.from_records(records)
    import pandas as pd
import numpy as np
import re

# --- helpers (safe to keep once) ---
def _find_anchor(df: pd.DataFrame, text: str):
    coords = np.where(df.astype(str).applymap(lambda x: isinstance(x, str) and text in x).values)
    if len(coords[0]) == 0:
        return None
    return int(coords[0][0]), int(coords[1][0])

def _is_number(x):
    return isinstance(x, (int, float, np.integer, np.floating)) and pd.notna(x)

# --- new extractor ---
def extract_company_pl_to_tidy(xlsx_path: str) -> pd.DataFrame:
    """
    Firm sheet -> Company Profit & Loss Statement to tidy format.
    Columns: section, metric, period, value, unit, sheet, row, col
    Handles Period 0..10 (or any count) and 'Cumulative'.
    """
    sheet = "Firm"
    unit = "thousands of dollars"
    df = pd.read_excel(xlsx_path, sheet_name=sheet, header=None)

    # 1) Find the anchor
    anchor = _find_anchor(df, "Company Profit & Loss Statement")
    if not anchor:
        raise ValueError("Company Profit & Loss Statement not found.")
    ar, ac = anchor

    # 2) Take a generous window under the anchor
    win = df.iloc[ar:ar+60, max(0, ac-1):ac+40].copy()
    base_c = max(0, ac-1)

    # 3) Find the header row that contains many period labels
    header_row = None
    period_cols = {}  # col index (in window) -> period label (int or "Cumulative")
    for r in range(win.shape[0]):
        tmp = {}
        for c in range(win.shape[1]):
            v = win.iat[r, c]
            if isinstance(v, str):
                s = v.strip()
                m = re.match(r"^Period\s*(\d+)$", s)
                if m:
                    tmp[c] = int(m.group(1))
                elif s.lower() == "cumulative":
                    tmp[c] = "Cumulative"
        if len(tmp) >= 3:  # enough columns to believe it is the header
            header_row = r
            period_cols = tmp
            break
    if header_row is None:
        raise ValueError("Could not locate the Period header row in P&L.")

    # 4) Walk metric rows below header until the block ends
    records = []
    for r in range(header_row + 1, win.shape[0]):
        row_texts = [win.iat[r, c] for c in range(win.shape[1])]
        # stop conditions
        joined = " ".join([str(x) for x in row_texts if isinstance(x, str)]).lower()
        if "market contribution" in joined:
            break
        if "all numbers" in joined:
            break
        if all(pd.isna(x) for x in row_texts):
            # blank row, but keep scanning a little more in case of spacing
            continue

        # find the leftmost text cell: that is the metric name
        metric = None
        for c in range(win.shape[1]):
            v = win.iat[r, c]
            if isinstance(v, str) and v.strip():
                metric = v.strip()
                break
        if not metric:
            continue

        # collect values for each period column
        for c, per_label in period_cols.items():
            if c >= win.shape[1]:
                continue
            val = win.iat[r, c]
            if _is_number(val):
                records.append({
                    "section": "Company Profit & Loss Statement",
                    "metric": metric,
                    "period": per_label,  # int or "Cumulative"
                    "value": float(val),
                    "unit": unit,
                    "sheet": sheet,
                    "row": ar + r,
                    "col": base_c + c,
                })

    return pd.DataFrame.from_records(records)

