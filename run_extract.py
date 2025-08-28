from markstrat_tidy_extractor import (
    extract_product_contribution_to_tidy,
    extract_company_pl_to_tidy,
)

excel_path = "TeamExport_A57286_Berlin_S_Period 7.xlsx"

pc = extract_product_contribution_to_tidy(excel_path)
pc.to_csv("product_contribution_tidy.csv", index=False)
print("Wrote product_contribution_tidy.csv", len(pc), "rows")

pl = extract_company_pl_to_tidy(excel_path)
pl.to_csv("firm_company_pl.csv", index=False)
print("Wrote firm_company_pl.csv", len(pl), "rows")
