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
de
