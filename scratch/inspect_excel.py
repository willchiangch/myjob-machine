import pandas as pd

file_path = r'd:\myproject\myjob-machine\issue_report_org\UAT測試回報清單_全聯會.xlsx'
df = pd.read_excel(file_path, nrows=1)
for i, col in enumerate(df.columns):
    # Print index and a cleaned version of the name to handle encoding
    print(f"Index {i}: {col}")
