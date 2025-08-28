import pandas as pd
from markstrat_tidy_extractor import extract_product_contribution_to_tidy

# Change this filename to match the Excel you uploaded
excel_path = "TeamExport_A57286_Berlin_S_Period 7.xlsx"

tidy = extract_product_contribution_to_tidy(excel_path)
tidy = tidy.sort_values(["product", "metric"]).reset_index(drop=True)
tidy.to_csv("product_contribution_tidy_latest.csv", index=False)

print("✅ Wrote product_contribution_tidy_latest.csv")
print(tidy.head(10))
