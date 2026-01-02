import pandas as pd
import glob
import os
import re

# 設定檔案目錄
folder_path = 'C:\\Users\\user\\workplace\\data\\'
ref_path = 'C:\\Users\\user\\workplace\\FB11407F通訊處20260101.xlsx'

# 標準化單位欄位的函式
def standardize_unit(val, mapping):
    if pd.isna(val) or not isinstance(val, str):
        return val
    
    # 1. 徹底消除所有空白 (暴力法)
    val = "".join(val.split())
    
    # 2. 檢查是否已經是正確格式 (前5碼英數 + 後面有中文)
    if re.match(r"^[a-zA-Z0-9]{5}.+", val):
        return val
    
    # 3. 如果 val 在 mapping 中 (只有代碼 或 只有名稱)
    if val in mapping:
        target = mapping[val]
        if re.match(r"^[a-zA-Z0-9]{5}", val): 
            return f"{val}{target}"
        else: # 如果輸入是中文
            return f"{target}{val}"
    
    # 4. 【進階查找】如果 val 是 "富宅TP838"，但 mapping 只有 "TP838"
    # 嘗試從字串中抽出 5 碼代碼來對照
    found_code = re.search(r"[a-zA-Z0-9]{5}", val)
    if found_code:
        code = found_code.group()
        if code in mapping:
            return f"{code}{mapping[code]}"
            
    return val

# 讀取代碼對照表
print("--- 讀取代碼對照表 ---")
try:
    ref_ = pd.read_excel(ref_path, skiprows=1) 
    ref_df = ref_[['代碼', '單位名稱']].copy()
    ref_df = ref_df.replace(['通訊處', '代碼', '單位名稱'], '', regex=True)
    ref_df = ref_df.replace(r'\s+', '', regex=True)
    ref_df['代碼'] = ref_df['代碼'].astype(str).str.strip()
    ref_df['單位名稱'] = ref_df['單位名稱'].astype(str).str.strip()
    ref_df = ref_df.dropna(subset=['單位名稱']) 
    ref_df = ref_df[ref_df['單位名稱'] != '']
    ref_df = ref_df[ref_df['單位名稱'] != 'nan']
    print(ref_df.head())
    print(f"代碼對照表讀取成功，共有 {len(ref_df)} 筆資料。")

    to_csv_path = 'C:\\Users\\user\\workplace\\ref_df.csv'
    ref_df.to_csv(to_csv_path, index=False, encoding='utf-8-sig')
    print(f"✅ 提取的資料已成功儲存至 '{to_csv_path}'。")

    # 建立雙向查找字典：
    # 一個是 代碼 -> 名稱，一個是 名稱 -> 代碼
    mapping_dict = dict(zip(ref_df['代碼'], ref_df['單位名稱']))
    mapping_dict.update(dict(zip(ref_df['單位名稱'], ref_df['代碼'])))
except FileNotFoundError:
    print("⚠️ 錯誤：找不到 'unit_mapping.csv' 檔案。請確保路徑正確。")
    exit()

# 讀取原始數據並提取所需欄位
print("--- 數據讀取與提取資料 ---")
# 抓取所有符合開頭為 "RFA" 且副檔名為 .csv 的檔案
file_list = glob.glob(os.path.join(folder_path, "RFA-*.csv"))

all_data_frames = []

for file in file_list:
    try:
        # 讀取單一檔案
        temp_df = pd.read_csv(file, skiprows=1, encoding='utf-8-sig')
        # 移除「序」欄位是空白的資料
        df = temp_df.dropna(subset=['序'])
        # 移除「連絡電話」欄位是空白的資料
        df = df.dropna(subset=['連絡電話'])  
        
        # 轉換「序」欄位為字串，並排除包含「取消」字樣的資料
        # 使用 .astype(str) 確保能處理所有格式，~ 代表「不包含」
        df = df[~df['序'].astype(str).str.contains('取消')]
        
        # 提取「單位」和「姓名」兩欄
        # 這裡建議使用 .copy() 避免 SettingWithCopyWarning
        extracted_data = df[['單位', '姓名']].copy()
        extracted_data = extracted_data.replace(r'\s+', '', regex=True).replace(['-', '一分處', '一', 'ㄧ', '分處'], '', regex=True)
        extracted_data['單位'] = extracted_data['單位'].str.upper()

        to_c = 'C:\\Users\\user\\workplace\\test.csv'
        extracted_data.to_csv(to_c, index=False, encoding='utf-8-sig')
        extracted_data['單位'] = extracted_data['單位'].apply(lambda x: standardize_unit(x, mapping_dict))
        all_data_frames.append(extracted_data)
        # 將提取的資料儲存為新的 CSV 檔案
        tocsv_path = 'C:\\Users\\user\\workplace\\extracted_data.csv'
        extracted_data.to_csv(tocsv_path, index=False, encoding='utf-8-sig')
        print(f"✅ 提取的資料已成功儲存至 '{tocsv_path}'。")
        
        print("原始數據讀取成功並進行提取。")
        print("-" * 30)
        print(extracted_data.head()) # 顯示前五筆檢查結果
        print("-" * 30)
        print(f"提取完成，共有 {len(extracted_data)} 筆有效資料。")

    except FileNotFoundError:
        print(f"⚠️ 錯誤：找不到 '{folder_path}' 檔案。請確保路徑正確。")
    except KeyError as e:
        print(f"⚠️ 錯誤：找不到欄位 {e}，請檢查 CSV 的欄位名稱是否正確。")
    except Exception as e:
        print(f"⚠️ 發生未知錯誤：{e}")

# 合併所有資料
if all_data_frames:
    extracted_data = pd.concat(all_data_frames, ignore_index=True)

# 統計各單位人數
# 使用 size() 統計行數，reset_index 轉換回 DataFrame 格式
summary_df = extracted_data.groupby('單位').size().reset_index(name='報名人數')

# 依照人數多寡排序 (由多到少)
# summary_df = summary_df.sort_values(by='報名人數', ascending=False)

print("--- 統計完成 ---")
print(summary_df)

# 匯出為 Excel 檔案
output_file = 'RFA報名人數統計表_2026.xlsx'

try:
    # 需要安裝 openpyxl: pip install openpyxl
    # index=False 代表不要把 pandas 的索引序號存進去

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='人數統計', index=False)
        extracted_data.to_excel(writer, sheet_name='詳細名單', index=False)

    # summary_df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"✅ 檔案已成功儲存為: {output_file}")
    
except Exception as e:
    print(f"❌ 儲存 Excel 時發生錯誤: {e}")
