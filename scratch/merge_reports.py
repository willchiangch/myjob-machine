import pandas as pd
import os
import glob
import re

# 設定路徑
input_dir = r'd:\myproject\myjob-machine\issue_report_org'
output_file = os.path.join(input_dir, '整合回報清單_2025.xlsx')

# 取得 5 個檔案
excel_files = glob.glob(os.path.join(input_dir, 'UAT測試回報清單_*.xlsx'))

# 定義模組關鍵字
module_keywords = {
    '會員與事務所': ['會員', '事務所'],
    '會費': ['費', '繳費', '扣款'],
    '公文': ['公文', '收文', '發文'],
    '助理': ['助理'],
    '活動與課程': ['活動', '課程', '報名']
}

# 定義性質關鍵字
nature_keywords = {
    'Bug': ['Bug', '錯誤', '異常', '失敗'],
    'CR': ['CR', '需求', '新增', '功能'],
    '優化': ['優化', '調整', '建議', '美化']
}

def get_module(desc):
    if pd.isna(desc): return '其他'
    for module, keywords in module_keywords.items():
        if any(kw in desc for kw in keywords):
            return module
    return '其他'

def get_nature(desc, n_val, o_val):
    # 1. N 欄有值 -> 調整DB
    if pd.notna(n_val):
        return '調整DB'
    
    # 2. O 欄有值 -> 依據 D 判斷 Bug/CR/優化
    if pd.notna(o_val):
        if pd.isna(desc): return '優化' 
        for nature, keywords in nature_keywords.items():
            if any(kw in desc for kw in keywords):
                return nature
        return '優化' 
    
    return '其他'

all_dfs = []

for file in excel_files:
    source_match = re.search(r'UAT測試回報清單_(.+)\.xlsx', os.path.basename(file))
    source = source_match.group(1) if source_match else '未知'
    
    print(f"正在處理: {source}...")
    
    df = pd.read_excel(file)
    
    # B 欄 (Index 1) 轉換為日期
    col_b_name = df.columns[1]
    df[col_b_name] = pd.to_datetime(df[col_b_name], errors='coerce')
    
    # 1. 時間篩選 (2025年)
    mask = (df[col_b_name] >= '2025-01-01') & (df[col_b_name] <= '2025-12-31')
    filtered_df = df.loc[mask].copy()
    
    if filtered_df.empty:
        continue

    result_rows = []
    for _, row in filtered_df.iterrows():
        desc = row.iloc[3]   # D 欄
        n_val = row.iloc[13] # N 欄
        o_val = row.iloc[14] # O 欄
        
        module = get_module(desc)
        nature = get_nature(desc, n_val, o_val)
        
        # 格式化日期為 yyyy/mm/dd 字串
        report_date = row.iloc[1]
        date_str = report_date.strftime('%Y/%m/%d') if pd.notna(report_date) else ''
        
        result_rows.append({
            '來源': source,
            'No.': row.iloc[0],
            '回報日期': date_str,      # 修改欄位名稱與格式
            '回報單位/人': row.iloc[2],
            '問題描述': row.iloc[3],
            '功能模組': module,        # 拆分欄位 1
            '問題性質': nature         # 拆分欄位 2
        })
    
    all_dfs.append(pd.DataFrame(result_rows))

if all_dfs:
    final_df = pd.concat(all_dfs, ignore_index=True)
    # 確保輸出順序
    cols = ['來源', 'No.', '回報日期', '回報單位/人', '問題描述', '功能模組', '問題性質']
    final_df = final_df[cols]
    final_df.to_excel(output_file, index=False)
    print(f"\n成功！整合檔案已儲存至: {output_file}")
    print(f"總筆數: {len(final_df)}")
else:
    print("\n警告：未找到符合條件的資料。")
