import streamlit as st
import pandas as pd
import os
import re
import io

# --- 1. 配置與核心邏輯 ---

def standardize_unit(val, mapping):
    if pd.isna(val) or not isinstance(val, str):
        return val
    val = "".join(val.split()).upper()
    
    # 正則匹配：前5碼英數 + 後面有內容
    if re.match(r"^[A-Z0-9]{5}.+", val):
        return val
    
    # 直接在字典中
    if val in mapping:
        target = mapping[val]
        if re.match(r"^[A-Z0-9]{5}$", val): 
            return f"{val}{target}"
        else:
            return f"{target}{val}"
    
    # 進階查找 (提取5碼代碼)
    found_code = re.search(r"[A-Z0-9]{5}", val)
    if found_code:
        code = found_code.group()
        if code in mapping:
            return f"{code}{mapping[code]}"
    return val

# --- 2. 讀取對照表 (關鍵修改：保留原始清單與順序) ---

@st.cache_data
def get_full_reference():
    try:
        # 建議維持讀取整張表，由邏輯來過濾
        ref_raw = pd.read_excel(REF_PATH, skiprows=0) 
        
        ref_list = []
        mapping = {}
        
        for _, row in ref_raw.iterrows():
            code = str(row['代碼']).strip().upper() if pd.notna(row['代碼']) else ""
            name = str(row['單位名稱']).strip() if pd.notna(row['單位名稱']) else ""
            
            # 過濾掉無意義的行
            if code in ["單位名稱", "NAN", "", "代碼"] and name in ["單位名稱", "NAN", "", "代碼"]:
                continue
            
            # --- 核心邏輯修正 ---
            # 只有當代碼長度剛好是 5 碼時，才判定為通訊處
            # 如果 code 其實是很長的一串字（如區部名稱），或是 name 是空的，就進入 else (標題模式)
            if len(code) == 5 and code != "NAN":
                # 這是真正的通訊處
                clean_name = name.replace("通訊處", "").replace("通訊", "")
                full_display = f"{code}{clean_name}"
                mapping[code] = clean_name
                mapping[clean_name] = code
                ref_list.append({"原始清單": full_display, "is_unit": True})
            else:
                # 進入這裡代表：code 是空的，或者是長串的標題文字
                # 我們優先取 name，如果 name 是空的，就取 code (因為標題可能跑去代碼欄)
                title_text = name if name not in ["", "NAN"] else code
                
                if title_text not in ["", "NAN"]:
                    short_name = title_text[:4] # 只取前四個字
                    ref_list.append({"原始清單": short_name, "is_unit": False})
            
        to_csv_path = 'C:\\Users\\user\\workplace\\RFA\\ref_df.csv'
        pd.DataFrame(ref_list).to_csv(to_csv_path, index=False, encoding='utf-8-sig')
        print(f"✅ 提取的資料已成功儲存至 '{to_csv_path}'。")
        
        return pd.DataFrame(ref_list), mapping
    except Exception as e:
        st.error(f"對照表讀取失敗：{e}")
        return pd.DataFrame(), {}

# --- 3. 處理數據 ---

def process_data(uploaded_file, mapping_dict):
    df = pd.read_csv(uploaded_file, skiprows=1, encoding='utf-8-sig')
    df = df.dropna(subset=['序', '連絡電話'])
    df = df[~df['序'].astype(str).str.contains('取消|轉班|轉讓', na=False)]
    
    extracted_data = df[['單位', '姓名']].copy()
    extracted_data = extracted_data.replace(r'\s+|-|一分處|一|ㄧ|分處|通訊', '', regex=True)
    extracted_data['單位'] = extracted_data['單位'].str.upper().apply(lambda x: standardize_unit(x, mapping_dict))

    tocsv_path = 'C:\\Users\\user\\workplace\\RFA\\extracted_data.csv'
    extracted_data.to_csv(tocsv_path, index=False, encoding='utf-8-sig')
    print(f"✅ 提取的資料已成功儲存至 '{tocsv_path}'。")
    
    return extracted_data

# --- 4. Streamlit 介面 ---

st.set_page_config(page_title="RFA 報名管理系統", layout="wide")
st.title("📊 RFA 報名資料增量更新系統 (完整架構版)")

MASTER_DB_PATH = 'master_data.csv'
REF_PATH = 'FB11407F通訊處20260101.xlsx'

# 獲取完整清單與字典
ref_df, mapping_dict = get_full_reference()

# 側邊欄與上傳邏輯 (與先前相同，略作精簡)
if os.path.exists(MASTER_DB_PATH):
    master_df = pd.read_csv(MASTER_DB_PATH)
    st.sidebar.success(f"🗃️ 資料庫筆數: {len(master_df)}")
else:
    master_df = pd.DataFrame(columns=['單位', '姓名'])

uploaded_files = st.file_uploader("上傳 RFA 報名 CSV", type="csv", accept_multiple_files=True)

if uploaded_files:
    new_dfs = [process_data(f, mapping_dict) for f in uploaded_files]
    current_batch = pd.concat(new_dfs, ignore_index=True)

    st.write("🔍 本次上傳預覽：")
    st.dataframe(current_batch.head(), use_container_width=True)

    if st.button("🚀 確認合併至主資料庫"):
        final_df = pd.concat([master_df, current_batch], ignore_index=True).drop_duplicates(subset=['單位', '姓名'], keep='last')
        final_df.to_csv(MASTER_DB_PATH, index=False, encoding='utf-8-sig')
        st.balloons()
        master_df = final_df

# --- 5. 統計與報表產出 (核心修改：Left Merge) ---

if not master_df.empty and not ref_df.empty:
    st.divider()
    
    # A. 算人數
    counts = master_df.groupby('單位').size().reset_index(name='報名人數')
    
    # B. 將人數併回完整清單 (用「原始清單」去對「單位」)
    # 這樣沒報名的單位會變成 NaN，標題列也會保留
    final_summary = pd.merge(ref_df, counts, left_on='原始清單', right_on='單位', how='left')
    
    # C. 清理結果：將單位的 NaN 轉為 0，但保持「標題列」的人數為空(比較美觀)
    final_summary['報名人數'] = final_summary.apply(
        lambda row: int(row['報名人數']) if pd.notna(row['報名人數']) 
        else (0 if row['is_unit'] else ""), axis=1
    )
    
    # 只留需要的欄位
    display_summary = final_summary[['原始清單', '報名人數']]

    st.subheader("第二步：數據統計 (依通訊錄順序)")
    col1, col2 = st.columns([2, 1])
    with col1:
        st.dataframe(display_summary, use_container_width=True, height=600)
    
    with col2:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            display_summary.to_excel(writer, sheet_name='人數統計(依順序)', index=False)
            master_df.to_excel(writer, sheet_name='詳細名單', index=False)
        
        
        st.download_button(
            label="📥 下載完整統計報表",
            data=buffer.getvalue(),
            file_name=f"RFA報名統計_{pd.Timestamp.now().strftime('%m%d')}.xlsx"
        )