import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

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
uploaded_files = st.file_uploader("請拖入 Invoice(.xls)、Packing(.xls) 與 客人檔案(.xlsx)", type=['xls', 'xlsx'], accept_multiple_files=True)

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
    if st.button("🚀 執行自動對位與回填", use_container_width=True):
        try:
            # 讀取函數：支援 .xls 與 .xlsx
            def smart_read(file_obj):
                engine = 'xlrd' if file_obj.name.lower().endswith('.xls') else 'openpyxl'
                return pd.read_excel(file_obj, engine=engine, dtype=str).fillna('')

            # A. 處理客人檔案 (計算重量與材積)
            df_cust = smart_read(files_dict["Customer"])
            cust_mapping = {}
            warning_list = []
            
            for _, row in df_cust.iterrows():
                delivery_id = str(row.iloc[2]).strip() # C 欄
                pcs_val = row.iloc[4] # E 欄
                
                if delivery_id.upper().startswith("GMJI"):
                    if not delivery_id.endswith("001"):
                        warning_list.append(delivery_id)
                    
                    try:
                        qty = float(pcs_val)
                        weight_f = round(qty * 0.1, 2)
                        cust_mapping[delivery_id] = {
                            "f": weight_f,
                            "g": round(weight_f - 0.05, 2),
                            "h": get_dimensions(qty)
                        }
                    except: pass

            if warning_list:
                st.error(f"⚠️ 偵測到 {len(warning_list)} 筆單號非 001 結尾！")
                st.warning(", ".join(warning_list))

            # B. 讀取原始 Packing 檔案內容 (無論 .xls 或 .xlsx)
            # 我們讀取全部內容，並重新建立一個 Workbook 保持格式一致
            df_pac_raw = smart_read(files_dict["Packing"])
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Updated_Packing"

            # 定義樣式
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            header_font = Font(name='Arial', size=10, bold=True)
            normal_font = Font(name='Arial', size=10)

            # 將原始數據寫入新表並進行比對修改
            sum_f = sum_g = 0.0
            data_start_row = 14
            match_count = 0
            
            # 寫入標題與原始資料 (模擬原始 Packing 結構)
            # 注意：DataFrame 索引從 0 開始，Excel 從 1 開始
            # 我們直接把原始 DataFrame 的內容填回去，但在第 14 行之後做邏輯判斷
            for r_idx, row_data in enumerate(df_pac_raw.values, 1):
                # 因為我們讀取時沒設 header，所以 r_idx 1 就是 Excel 的第 1 行
                for c_idx, value in enumerate(row_data, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=value)
                    cell.font = normal_font
                    
                    # 邏輯判斷：如果是第 14 行以後，且是在處理 C 欄 (單號)
                    if r_idx >= data_start_row:
                        p_id = str(df_pac_raw.iloc[r_idx-1, 2]).strip() # 原始 C 欄單號
                        
                        if p_id in cust_mapping:
                            data = cust_mapping[p_id]
                            # 修改 F(6), G(7), H(8) 欄
                            if c_idx == 6: 
                                cell.value = data["f"]
                                if r_idx == data_start_row: # 僅在第一次比對到時開始累加
                                     pass 
                            elif c_idx == 7: cell.value = data["g"]
                            elif c_idx == 8: cell.value = data["h"]
                            
                            # 為了最後加總，我們在處理該列最後一個欄位時累加數值
                            if c_idx == 1: # 每一列只加一次
                                sum_f += data["f"]
                                sum_g += data["g"]
                                match_count += 1

            # 找到最後一筆資料列
            final_data_row = len(df_pac_raw)
            total_row = final_data_row + 1
            
            # 填入加總
            ws.cell(row=total_row, column=5, value="TOTAL:").font = header_font
            ws.cell(row=total_row, column=6, value=round(sum_f, 2)).font = header_font
            ws.cell(row=total_row, column=7, value=round(sum_g, 2)).font = header_font

            # C. 產出檔案
            output = BytesIO()
            wb.save(output)
            st.balloons()
            st.success(f"✅ 原始 .xls 檔案處理完成！共比對 {match_count} 筆。")
            
            st.download_button(
                label="📥 下載更新後的全球報關 Packing",
                data=output.getvalue(),
                file_name="Global_Packing_Final.xlsx",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"轉換異常：{e}")
