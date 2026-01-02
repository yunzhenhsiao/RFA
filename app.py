import streamlit as st
import pandas as pd
import os
import re
import io

# --- 1. 配置與核心邏輯 (整合 data_merge.py 的所有功能) ---

def standardize_unit(val, mapping):
    """標準化單位欄位，支援進階查找與暴力去空白"""
    if pd.isna(val) or not isinstance(val, str):
        return val
    
    # 1. 徹底消除所有空白 (暴力法) - 來自 data_merge.py
    val = "".join(val.split())
    
    # 2. 轉大寫 (確保一致性)
    val = val.upper()
    
    # 3. 檢查是否已經是正確格式 (前5碼英數 + 後面有中文內容) - 來自 data_merge.py
    # 原本 app.py 是寫 [a-zA-Z]{2,}\d{3}，這裡改用 data_merge 的通用格式 [a-zA-Z0-9]{5}.+
    if re.match(r"^[A-Z0-9]{5}.+", val):
        return val
    
    # 4. 如果 val 直接在 mapping 中 (只有代碼 或 只有名稱)
    if val in mapping:
        target = mapping[val]
        # 如果輸入是 5 碼代碼
        if re.match(r"^[A-Z0-9]{5}$", val): 
            return f"{val}{target}"
        else: # 如果輸入是純中文名稱
            return f"{target}{val}"
    
    # 5. 【進階查找】 - 來自 data_merge.py
    # 嘗試從字串中抽出 5 碼代碼來對照 (例如輸入 "富宅TP838" -> 提取 "TP838")
    found_code = re.search(r"[A-Z0-9]{5}", val)
    if found_code:
        code = found_code.group()
        if code in mapping:
            return f"{code}{mapping[code]}"
            
    return val

def process_data(uploaded_file, mapping_dict):
    """處理單一上傳檔案的清理流程 (整合 data_merge 的過濾條件)"""
    # 讀取檔案，跳過第一列 (skiprows=1)
    df = pd.read_csv(uploaded_file, skiprows=1, encoding='utf-8-sig')
    
    # 移除「序」與「連絡電話」空白的資料
    df = df.dropna(subset=['序', '連絡電話'])
    
    # 排除包含「取消」字樣的資料
    df = df[~df['序'].astype(str).str.contains('取消')]
    
    # 提取「單位」和「姓名」
    extracted_data = df[['單位', '姓名']].copy()
    
    # 清理字串內容 - 整合了 data_merge.py 的 replace 規則 (包含 -, 一分處, ㄧ, 分處等)
    extracted_data = extracted_data.replace(r'\s+', '', regex=True)
    extracted_data = extracted_data.replace(['-', '一分處', '一', 'ㄧ', '分處'], '', regex=True)
    
    # 統一轉大寫並執行標準化
    extracted_data['單位'] = extracted_data['單位'].str.upper()
    extracted_data['單位'] = extracted_data['單位'].apply(lambda x: standardize_unit(x, mapping_dict))
    
    return extracted_data

# --- 2. Streamlit 網頁介面 ---

st.set_page_config(page_title="RFA 報名管理系統", layout="wide")
st.title("📊 RFA 報名資料增量更新系統")

# 設定路徑 (使用你 data_merge.py 中的 Excel 路徑)
MASTER_DB_PATH = 'master_data.csv'
REF_PATH = 'FB11407F通訊處20260101.xlsx'

# 讀取對照表 (整合 data_merge.py 的 Excel 清洗邏輯)
@st.cache_data
def get_mapping():
    try:
        # 讀取 Excel 並套用 data_merge 的清洗流程
        ref_raw = pd.read_excel(REF_PATH, skiprows=1) 
        ref_df = ref_raw[['代碼', '單位名稱']].copy()
        
        # 移除標題字眼與空白
        ref_df = ref_df.replace(['通訊處', '代碼', '單位名稱'], '', regex=True)
        ref_df = ref_df.replace(r'\s+', '', regex=True)
        
        # 欄位格式化
        ref_df['代碼'] = ref_df['代碼'].astype(str).str.strip().str.upper()
        ref_df['單位名稱'] = ref_df['單位名稱'].astype(str).str.strip()
        
        # 移除空值與無效字串 (nan)
        ref_df = ref_df.dropna(subset=['單位名稱']) 
        ref_df = ref_df[~ref_df['單位名稱'].isin(['', 'nan'])]
        
        # 建立雙向字典
        m = dict(zip(ref_df['代碼'], ref_df['單位名稱']))
        m.update(dict(zip(ref_df['單位名稱'], ref_df['代碼'])))
        return m
    except Exception as e:
        st.error(f"⚠️ 對照表讀取失敗，請確認路徑：{REF_PATH}")
        st.error(f"錯誤訊息: {e}")
        return {}

mapping_dict = get_mapping()

# 側邊欄：顯示當前主資料庫狀態
if os.path.exists(MASTER_DB_PATH):
    # 強制讀取為字串避免 ID 被科學符號化
    master_df = pd.read_csv(MASTER_DB_PATH)
    st.sidebar.success(f"🗃️ 目前資料庫已有: {len(master_df)} 筆資料")
else:
    master_df = pd.DataFrame(columns=['單位', '姓名'])
    st.sidebar.info("📂 目前資料庫為空")

# --- 3. 檔案上傳區 ---
st.subheader("第一步：上傳新資料")
uploaded_files = st.file_uploader("選擇 RFA 報名 CSV 檔案 (支援多選)", type="csv", accept_multiple_files=True)

if uploaded_files:
    all_new_frames = []
    for f in uploaded_files:
        temp_df = process_data(f, mapping_dict)
        all_new_frames.append(temp_df)
    
    current_batch_df = pd.concat(all_new_frames, ignore_index=True)
    
    st.write("🔍 本次上傳預覽：")
    st.dataframe(current_batch_df.head(), use_container_width=True)

    # --- 4. 增量更新按鈕 ---
    if st.button("🚀 確認合併至主資料庫"):
        # 合併舊資料與新資料
        # 以「單位+姓名」作為唯一基準避免重複重複
        final_df = pd.concat([master_df, current_batch_df], ignore_index=True)
        final_df = final_df.drop_duplicates(subset=['單位', '姓名'], keep='last')
        
        final_df.to_csv(MASTER_DB_PATH, index=False, encoding='utf-8-sig')
        st.balloons()
        st.success(f"✅ 更新成功！目前總數：{len(final_df)} 筆。")
        master_df = final_df # 即時更新變數供下方統計顯示

# --- 5. 統計與下載區 ---
if not master_df.empty:
    st.divider()
    st.subheader("第二步：數據統計與下載")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # 統計各單位人數
        summary_df = master_df.groupby('單位').size().reset_index(name='報名人數')
        # 依人數降冪排序
        summary_df = summary_df.sort_values(by='報名人數', ascending=False)
        st.dataframe(summary_df, use_container_width=True)
    
    with col2:
        # 產出 Excel 下載流
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            summary_df.to_excel(writer, sheet_name='人數統計', index=False)
            master_df.to_excel(writer, sheet_name='詳細名單', index=False)
        
        st.download_button(
            label="📥 下載完整統計 Excel 報表",
            data=buffer.getvalue(),
            file_name=f"RFA報名統計表_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )