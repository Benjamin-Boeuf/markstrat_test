from markstrat_tidy_extractor import (
    extract_product_contribution_to_tidy,
    extract_company_pl_to_tidy,
)

excel_path = "TeamExport_A57286_Berlin_S_Period 7.xlsx"

pc = extract_product_contribution_to_tidy(excel_path)
pc = pc.sort_values(["product", "metric"]).reset_index(drop=True)
pc.to_csv("product_contribution_tidy.csv", index=False)
print("Wrote product_contribution_tidy.csv", len(pc), "rows")

pl = extract_company_pl_to_tidy(excel_path)
# sort with cumulative last
def _per_sort_key(x):
    return 999 if x == "Cumulative" else int(x)
pl = pl.sort_values(["metric","period"], key=lambda s: s.map(_per_sort_key)).reset_index(drop=True)
pl.to_csv("firm_company_pl.csv", index=False)
print("Wrote firm_company_pl.csv", len(pl), "rows")
