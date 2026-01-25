import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# 1. 網頁基本設定
st.set_page_config(page_title="全球報關文件轉換器", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #0E1117; }
    .big-title { font-size: 30px !important; font-weight: bold; color: #FFFFFF !important; }
    .stFileUploader section { background-color: #FFFFFF !important; border-radius: 10px; }
    div.stButton > button { background-color: #FFFFFF !important; color: #000000 !important; border: 2px solid #000000 !important; height: 50px; font-weight: bold; width: 100%; }
    .stMarkdown p, label { color: #FFFFFF !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="big-title">🌍 全球報關文件轉換器 (SG/MY)</p>', unsafe_allow_html=True)

# 2. 檔案上傳
uploaded_files = st.file_uploader("請拖入 Invoice、Packing 與 客人檔案", type=['xls', 'xlsx'], accept_multiple_files=True)

files_dict = {"Invoice": None, "Packing": None, "Customer": None}

if uploaded_files:
    for f in uploaded_files:
        fname = f.name.lower()
        if "invoice" in fname: files_dict["Invoice"] = f
        elif "packing" in fname: files_dict["Packing"] = f
        elif "pcs" in fname or "好馬吉" in fname: files_dict["Customer"] = f

# 3. 材積計算邏輯
def get_dimensions(pcs):
    try:
        val = int(float(pcs))
        if 1 <= val <= 5: return "18*19*15"
        elif 6 <= val <= 10: return "23*14*14"
        elif 11 <= val <= 40: return "29*20*20"
        elif val >= 41: return "42*30*25"
        return ""
    except: return ""

# 4. 轉換邏輯
if all(files_dict.values()):
    if st.button("🚀 開始全球報關資料核對與產出", use_container_width=True):
        try:
            # --- 自動偵測引擎讀取函數 ---
            def safe_read(file_obj):
                # 根據副檔名選擇引擎
                engine = 'xlrd' if file_obj.name.lower().endswith('.xls') else 'openpyxl'
                return pd.read_excel(file_obj, engine=engine, dtype=str).fillna('0')

            # 讀取數據
            df_cust = safe_read(files_dict["Customer"])
            cust_data = {}
            warning_list = []
            
            for _, row in df_cust.iterrows():
                # 抓取 C 欄 (index 2) 與 E 欄 (index 4)
                delivery_id = str(row.iloc[2]).strip()
                pcs = row.iloc[4]
                
                if delivery_id.upper().startswith("GMJI"):
                    # 尾數 001 檢查
                    if not delivery_id.endswith("001"): 
                        warning_list.append(delivery_id)
                    
                    try:
                        qty = float(pcs)
                        weight_f = qty * 0.1
                        cust_data[delivery_id] = {
                            "f": weight_f, 
                            "g": weight_f - 0.05, 
                            "h": get_dimensions(qty)
                        }
                    except: pass

            # 顯示單號警告
            if warning_list:
                st.error(f"⚠️ 偵測到 {len(warning_list)} 筆單號非 001 結尾：")
                st.warning(", ".join(warning_list))

            # --- 處理目標 Packing 檔案 ---
            # openpyxl 僅支援修改 .xlsx 檔案
            if not files_dict["Packing"].name.lower().endswith('.xlsx'):
                st.error("❌ 格式不符：Packing 檔案必須是 .xlsx 格式才能直接回填。")
                st.info("請將您的 Packing 打開並『另存新檔』為 Excel 活頁簿 (.xlsx) 後重新上傳。")
                st.stop()

            # 使用 openpyxl 載入範本
            wb = load_workbook(files_dict["Packing"])
            ws = wb.active
            sum_f = sum_g = 0.0
            last_row = 14
            match_count = 0

            # 從第 14 行開始尋找 C 欄單號
            for row_idx in range(14, ws.max_row + 1):
                p_id = str(ws.cell(row=row_idx, column=3).value).strip()
                if p_id in cust_data:
                    data = cust_data[p_id]
                    # 回填 F, G, H 欄
                    ws.cell(row=row_idx, column=6, value=data["f"])
                    ws.cell(row=row_idx, column=7, value=data["g"])
                    ws.cell(row=row_idx, column=8, value=data["h"])
                    sum_f += data["f"]
                    sum_g += data["g"]
                    match_count += 1
                    last_row = row_idx

            # 插入 TOTAL 加總列
            total_row = last_row + 1
            ws.cell(row=total_row, column=5, value="TOTAL:").font = Font(bold=True, name='Arial')
            ws.cell(row=total_row, column=6, value=round(sum_f, 2)).font = Font(bold=True, name='Arial')
            ws.cell(row=total_row, column=7, value=round(sum_g, 2)).font = Font(bold=True, name='Arial')

            # 轉為二進位流供下載
            output = BytesIO()
            wb.save(output)
            st.balloons()
            st.success(f"✅ 核對完成！成功比對回填 {match_count} 筆單號。")
            st.download_button(
                label="📥 下載更新後的 Packing (.xlsx)", 
                data=output.getvalue(), 
                file_name="Global_Packing_Updated.xlsx", 
                use_container_width=True
            )
        except Exception as e:
            st.error(f"系統發生異常：{e}")
