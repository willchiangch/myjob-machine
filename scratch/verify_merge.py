import pandas as pd

output_file = r'd:\myproject\myjob-machine\issue_report_org\整合回報清單_2025.xlsx'
df = pd.read_excel(output_file)

print("Columns:", df.columns.tolist())
print("\nFirst 5 rows:")
print(df.head(5).to_string())
