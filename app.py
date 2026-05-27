import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
import math

# 1. 網頁基本設定
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

st.markdown('<p class="big-title">🌍 全球 Packing 轉換器 (終極精確版)</p>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("請上傳您的原始 Packing 檔案 (.xls)", type=['xls', 'xlsx'])

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
            with st.spinner("正在設定活頁簿底層字體並計算精確排版..."):
                engine = 'xlrd' if uploaded_file.name.lower().endswith('.xls') else 'openpyxl'
                df = pd.read_excel(uploaded_file, header=None, dtype=str).fillna('')

                # --- A. 擷取上方表頭純文字 ---
                header_rows_clean = []
                for r_idx in range(11):
                    row_vals = [str(x).strip() for x in df.iloc[r_idx].tolist() if str(x).strip() != ""]
                    header_rows_clean.append(row_vals)

                # --- B. 擷取並過濾商品資料 ---
                items = []
                for idx in range(11, len(df)):
                    if "總箱數" in str(df.iloc[idx, 0]) or "TOTAL" in str(df.iloc[idx, 0]).upper():
                        break
                    desc_zh = str(df.iloc[idx, 2]).strip()
                    if "SHIPFEE" in desc_zh.upper():
                        continue
                    items.append(df.iloc[idx].tolist())

                # --- C. 群組化 SKU 並計算 ---
                processed_groups = []
                current_group = []
                for row in items[1:]:
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

                # --- D. 建立新 Excel 檔案 (底層字體修改) ---
                wb = Workbook()
                ws = wb.active
                ws.title = "Processed_Packing"

                global_font = Font(name='Calibri', size=12, bold=True)
                
                # 【本次修復點】：加上底線，正確讀取 NamedStyle 物件！
                for style in wb._named_styles:
                    if style.name == 'Normal':
                        style.font = global_font

                thin_border = Border(left=Side(style='thin', color='000000'), 
                                     right=Side(style='thin', color='000000'), 
                                     top=Side(style='thin', color='000000'), 
                                     bottom=Side(style='thin', color='000000'))

                # 1. 填入上方表頭文字
                for r_idx, row_vals in enumerate(header_rows_clean, 1):
                    ws.row_dimensions[r_idx].height = 15 # 預設高度
                    if r_idx in [1, 2, 4, 6]:
                        if len(row_vals) >= 1: ws.cell(row=r_idx, column=1, value=row_vals[0])
                    elif r_idx in [3, 5, 7, 8, 10]:
                        if len(row_vals) >= 1: ws.cell(row=r_idx, column=1, value=row_vals[0])
                        if len(row_vals) >= 2: ws.cell(row=r_idx, column=4, value=row_vals[1])
                    elif r_idx == 9:
                        if len(row_vals) >= 1: ws.cell(row=9, column=1, value=row_vals[0])
                        if len(row_vals) >= 2: ws.cell(row=9, column=3, value=row_vals[1])
                        if len(row_vals) >= 3: ws.cell(row=9, column=6, value=row_vals[2])

                # D7~F7 與 A8~C8 的動態列高計算 (防止文字被擋住)
                d7_val = str(ws.cell(row=7, column=4).value or "")
                if d7_val:
                    lines_d7 = math.ceil(len(d7_val) / 25) + d7_val.count('\n')
                    ws.row_dimensions[7].height = max(15, lines_d7 * 16.5)

                a8_val = str(ws.cell(row=8, column=1).value or "")
                if a8_val:
                    lines_a8 = math.ceil(len(a8_val) / 30) + a8_val.count('\n')
                    ws.row_dimensions[8].height = max(15, lines_a8 * 16.5)

                # 2. 寫入第 12 行標題
                ws.row_dimensions[12].height = 15
                headers_row12 = ["SKU", "Description_of_Goods_(zh)", "Qty", "Net_Weight_(KG)", "Gross_Weight_(KG)", "Measurement_(cm)"]
                for c_idx, val in enumerate(headers_row12, 1):
                    cell = ws.cell(row=12, column=c_idx, value=val)

                # 3. 寫入商品資料列 (從13行開始)
                start_row = 13
                for r_idx, row_data in enumerate(final_output_rows, start_row):
                    max_lines = 1
                    for c_idx, val in enumerate(row_data, 1):
                        cell = ws.cell(row=r_idx, column=c_idx, value=val)
                        
                        # 品名(B欄)自動換行與精確列高計算
                        if c_idx == 2 and val != "":
                            text_str = str(val)
                            # B欄寬 37.09，Calibri 12粗體約容納 15 個中文字
                            chars_per_line = 15
                            lines_by_length = math.ceil(len(text_str) / chars_per_line)
                            lines_by_enter = text_str.count('\n') + 1
                            max_lines = max(max_lines, lines_by_length, lines_by_enter)
                        
                        if val != "" and c_idx in [3, 4, 5]:
                            try: cell.value = float(val)
                            except: pass

                    # 動態安全高度計算
                    ws.row_dimensions[r_idx].height = max(15, max_lines * 16.5)

                # 4. 寫入最底部的總計列
                total_row_idx = start_row + len(final_output_rows)
                ws.row_dimensions[total_row_idx].height = 15
                qty_display = int(total_qty_sum) if total_qty_sum.is_integer() else round(total_qty_sum, 2)
                total_data = [f"總箱數:{sku_count}箱", "", qty_display, total_net_sum, total_gross_sum, ""]
                for c_idx, val in enumerate(total_data, 1):
                    ws.cell(row=total_row_idx, column=c_idx, value=val)

                # --- E. 完美合併與對齊配置 ---
                # 1. 表頭合併
                merge_rules = [
                    ('A1:F1', 'center'), ('A2:F2', 'center'), ('A4:F4', 'center'),
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
                
                # F9 單獨靠左
                ws['F9'].alignment = Alignment(horizontal='left', vertical='center')

                # 特別指定 D7 與 A8 開啟自動換行
                ws['D7'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                ws['A8'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

                # 2. 表格區塊對齊 (A~C置左，D~F置中)
                for r_idx in range(12, total_row_idx + 1):
                    for c_idx in range(1, 7):
                        cell = ws.cell(row=r_idx, column=c_idx)
                        if c_idx <= 3:
                            if c_idx == 2:
                                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                            else:
                                cell.alignment = Alignment(horizontal='left', vertical='center')
                        else:
                            cell.alignment = Alignment(horizontal='center', vertical='center')

                # 3. 劃格線與格式化
                for r_idx in range(12, total_row_idx + 1):
                    for c_idx in range(1, 7):
                        cell = ws.cell(row=r_idx, column=c_idx)
                        cell.border = thin_border
                        # 強制字體再次套用，以防萬一
                        cell.font = global_font
                        if r_idx >= 12:
                            if c_idx == 3 and isinstance(cell.value, (int, float)):
                                cell.number_format = '0'
                            elif c_idx in [4, 5] and isinstance(cell.value, (int, float)):
                                cell.number_format = '0.00'

                # 確保表頭的字體也都套用
                for r in range(1, 12):
                    for c in range(1, 7):
                        ws.cell(row=r, column=c).font = global_font

                # --- F. 設定精準自訂欄寬 ---
                col_widths = {'A': 18.82, 'B': 37.09, 'C': 3.91, 'D': 15.64, 'E': 17.18, 'F': 17.55}
                for col, width in col_widths.items():
                    ws.column_dimensions[col].width = width

                # 輸出
                output = BytesIO()
                wb.save(output)
                st.balloons()
                st.success("✅ 全球 Packing 轉換成功！(當機問題已修正，欄寬與換行已完美套用)")
                st.download_button(
                    label="📥 下載精確排版 Packing (.xlsx)",
                    data=output.getvalue(),
                    file_name="Global_Packing_Perfect.xlsx",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"系統轉換發生錯誤：{e}")
