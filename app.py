import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
import math

# 1. 網頁基本設定 (維持黑底高質感風格)
st.set_page_config(page_title="全球 Packing 轉換器", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #0E1117; }
    .big-title { font-size: 30px !important; font-weight: bold; color: #FFFFFF !important; }
    .stFileUploader section { background-color: #FFFFFF !important; border-radius: 10px; }
    div.stButton > button { background-color: #FFFFFF !important; color: #000000 !important; border: 2px solid #000000 !important; height: 50px; font-weight: bold; width: 100%; }
    .stMarkdown p, label { color: #FFFFFF !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="big-title">🌍 全球 Packing 轉換器 (Calibri 自訂版)</p>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("請上傳您的原始 Packing 檔案 (.xls)", type=['xls', 'xlsx'])

# 材積自動判別區間邏輯
def get_dimensions(pcs):
    try:
        val = int(float(pcs))
        if 1 <= val <= 5: return "18*19*15"
        elif 6 <= val <= 10: return "23*14*14"
        elif 11 <= val <= 40: return "29*20*20"
        elif val >= 41: return "42*30*25"
        return ""
    except: return ""

if uploaded_file:
    if st.button("🚀 執行智能規格轉換", use_container_width=True):
        try:
            with st.spinner("正在依照全新自訂規格進行排版與轉換..."):
                engine = 'xlrd' if uploaded_file.name.lower().endswith('.xls') else 'openpyxl'
                df = pd.read_excel(uploaded_file, header=None, dtype=str).fillna('')

                # --- A. 擷取上方表頭 (第 1 到 11 行) ---
                header_data = df.iloc[0:11].values.tolist()

                # --- B. 擷取並過濾下方商品資料 ---
                items = []
                for idx in range(11, len(df)):
                    if "總箱數" in str(df.iloc[idx, 0]) or "TOTAL" in str(df.iloc[idx, 0]).upper():
                        break
                    
                    desc_zh = str(df.iloc[idx, 2]).strip()
                    # 剔除 SHIPFEE 運費
                    if "SHIPFEE" in desc_zh.upper():
                        continue
                    
                    items.append(df.iloc[idx].tolist())

                # --- C. 群組化 SKU 並計算重量/材積 ---
                processed_groups = []
                current_group = []

                for row in items[1:]: # 跳過原本第12行的舊標題
                    sku = str(row[0]).strip()
                    if sku != "":
                        if current_group:
                            processed_groups.append(current_group)
                        current_group = [row]
                    else:
                        current_group.append(row)
                if current_group:
                    processed_groups.append(current_group)

                final_output_rows = []
                total_qty_sum = 0.0
                total_net_sum = 0.0
                total_gross_sum = 0.0
                sku_count = len(processed_groups)

                for group in processed_groups:
                    group_qty = 0.0
                    for r in group:
                        try: group_qty += float(r[3])
                        except: pass
                    
                    total_qty_sum += group_qty
                    gross_w = round(group_qty * 0.1, 2)
                    net_w = round(gross_w - 0.05, 2)
                    if net_w < 0: net_w = 0.0
                    
                    total_gross_sum += gross_w
                    total_net_sum += net_w
                    meas_str = get_dimensions(group_qty)
                    
                    for i, r in enumerate(group):
                        desc = str(r[2])
                        for kw in ["【新品】", "★", "【歡樂智多星推薦】"]:
                            desc = desc.replace(kw, "")
                        desc = desc.strip()
                        
                        out_sku = r[0] if i == 0 else ""
                        out_qty = r[3]
                        out_net = net_w if i == 0 else ""
                        out_gross = gross_w if i == 0 else ""
                        out_meas = meas_str if i == 0 else ""
                        
                        final_output_rows.append([out_sku, desc, out_qty, out_net, out_gross, out_meas])

                # --- D. 建立全新的 Excel 檔案並配置樣式 ---
                wb = Workbook()
                ws = wb.active
                ws.title = "Processed_Packing"

                # 全新字體規格：Calibri, 12號, 粗體
                global_font = Font(name='Calibri', size=12, bold=True)
                
                # 表格格線設定 (細實線)
                thin_border = Border(left=Side(style='thin'), 
                                     right=Side(style='thin'), 
                                     top=Side(style='thin'), 
                                     bottom=Side(style='thin'))

                # 1. 寫入上方表頭 (1~11行)
                for r_idx, r_data in enumerate(header_data, 1):
                    ws.row_dimensions[r_idx].height = 15 # 預設列高 15
                    for c_idx, val in enumerate(r_data, 1):
                        if c_idx <= 8:
                            cell = ws.cell(row=r_idx, column=c_idx, value=val)
                            cell.font = global_font
                            cell.alignment = Alignment(vertical='center')

                # 2. 寫入第 12 行全新標題
                ws.row_dimensions[12].height = 15
                headers_row12 = ["SKU", "Description_of_Goods_(zh)", "Qty", "Net_Weight_(KG)", "Gross_Weight_(KG)", "Measurement_(cm)"]
                for c_idx, val in enumerate(headers_row12, 1):
                    cell = ws.cell(row=12, column=c_idx, value=val)
                    cell.font = global_font
                    cell.border = thin_border
                    # A12~C12 置左，D12~F12 置中
                    if c_idx <= 3:
                        cell.alignment = Alignment(horizontal='left', vertical='center')
                    else:
                        cell.alignment = Alignment(horizontal='center', vertical='center')

                # 3. 寫入商品資料列 (13行開始)
                start_row = 13
                for r_idx, row_data in enumerate(final_output_rows, start_row):
                    max_lines = 1
                    
                    for c_idx, val in enumerate(row_data, 1):
                        cell = ws.cell(row=r_idx, column=c_idx, value=val)
                        cell.font = global_font
                        cell.border = thin_border
                        
                        # A~C欄 置左 (品名B欄開啟自動換行與動態列高計算)
                        if c_idx <= 3:
                            if c_idx == 2 and val != "":
                                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
                                text_str = str(val)
                                # Calibri 12號粗體在欄寬37.09下，約可穩當容納 15 個中文字
                                chars_per_line = 15 
                                lines_by_length = math.ceil(len(text_str) / chars_per_line)
                                lines_by_enter = text_str.count('\n') + 1
                                max_lines = max(max_lines, lines_by_length, lines_by_enter)
                            else:
                                cell.alignment = Alignment(horizontal='left', vertical='center')
                                if c_idx == 3 and val != "":
                                    try:
                                        cell.value = float(val)
                                        cell.number_format = '0'
                                    except: pass
                        # D~F欄 置中
                        else:
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            if val != "" and c_idx in [4, 5]:
                                try:
                                    cell.value = float(val)
                                    cell.number_format = '0.00'
                                except: pass

                    # 動態調整列高：最少 15，多行時依比例安全撐大
                    ws.row_dimensions[r_idx].height = max(15, max_lines * 16.5)

                # 4. 寫入最底部的總計列
                total_row_idx = start_row + len(final_output_rows)
                ws.row_dimensions[total_row_idx].height = 15
                qty_display = int(total_qty_sum) if total_qty_sum.is_integer() else round(total_qty_sum, 2)
                
                total_data = [f"總箱數:{sku_count}箱", "", qty_display, total_net_sum, total_gross_sum, ""]
                for c_idx, val in enumerate(total_data, 1):
                    cell = ws.cell(row=total_row_idx, column=c_idx, value=val)
                    cell.font = global_font
                    cell.border = thin_border
                    
                    # 遵循 A~C置左、D~F置中
                    if c_idx <= 3:
                        cell.alignment = Alignment(horizontal='left', vertical='center')
                        if c_idx == 3 and isinstance(val, (int, float)):
                            cell.value = float(val)
                            cell.number_format = '0'
                    else:
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        if val != "" and c_idx in [4, 5]:
                            try:
                                cell.value = float(val)
                                cell.number_format = '0.00'
                            except: pass

                # --- E. 完美配置指定之合併儲存格與對齊 ---
                merge_rules = [
                    # 合併置中
                    ('A1:F1', 'center'), ('A2:F2', 'center'), ('A4:F4', 'center'),
                    # 合併至右 (依據前版靠左排版)
                    ('A3:C3', 'left'), ('D3:F3', 'left'),
                    ('A5:C5', 'left'), ('D5:F5', 'left'),
                    ('A6:F6', 'left'),
                    ('A7:C7', 'left'), ('D7:F7', 'left'), 
                    ('A8:C8', 'left'), ('D8:F8', 'left'), 
                    ('A9:B9', 'left'), ('C9:E9', 'left'), 
                    ('A10:C10', 'left'), ('D10:F10', 'left')
                ]
                
                for m_range, align in merge_rules:
                    try:
                        ws.merge_cells(m_range)
                        top_left = m_range.split(':')[0]
                        ws[top_left].alignment = Alignment(horizontal=align, vertical='center')
                    except: pass 

                # --- F. 設定您指定的精確欄寬 ---
                col_widths = {'A': 18.82, 'B': 37.09, 'C': 3.91, 'D': 15.64, 'E': 17.18, 'F': 17.55}
                for col, width in col_widths.items():
                    ws.column_dimensions[col].width = width

                # 輸出
                output = BytesIO()
                wb.save(output)
                st.balloons()
                st.success("✅ 全球 Packing 轉換成功！已完全套用 Calibri 12號粗體與精確對齊。")
                st.download_button(
                    label="📥 下載全新自訂排版 Packing (.xlsx)",
                    data=output.getvalue(),
                    file_name="Global_Packing_Calibri_Custom.xlsx",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"系統轉換發生錯誤：{e}")
