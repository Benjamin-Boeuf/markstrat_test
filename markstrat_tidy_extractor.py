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
