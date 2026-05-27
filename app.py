import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Alignment

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

st.markdown('<p class="big-title">🌍 全球 Packing 轉換器 (完美排版版)</p>', unsafe_allow_html=True)

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
            with st.spinner("正在執行精準排版與資料重組..."):
                engine = 'xlrd' if uploaded_file.name.lower().endswith('.xls') else 'openpyxl'
                df = pd.read_excel(uploaded_file, header=None, dtype=str).fillna('')

                # --- A. 擷取上方表頭 (第 1 到 11 行) ---
                header_data = df.iloc[0:11].values.tolist()

                # --- B. 擷取並過濾下方商品資料 ---
                items = []
                for idx in range(11, len(df)):
                    # 遇到總計列就停止
                    if "總箱數" in str(df.iloc[idx, 0]) or "TOTAL" in str(df.iloc[idx, 0]).upper():
                        break
                    
                    desc_zh = str(df.iloc[idx, 2]).strip()
                    # 殺手級清理：遇到 SHIPFEE 直接跳過不加入清單
                    if "SHIPFEE" in desc_zh.upper():
                        continue
                    
                    items.append(df.iloc[idx].tolist())

                # --- C. 群組化 SKU 並計算重量/材積 ---
                processed_groups = []
                current_group = []

                for row in items[1:]: # [1:] 是為了跳過原本第12行的標題
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
                    
                    # 重組每一行，並剔除英文品名與 U/M
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
                        
                        # 新欄位順序: A(SKU), B(中文品名), C(Qty), D(Net), E(Gross), F(Meas)
                        final_output_rows.append([out_sku, desc, out_qty, out_net, out_gross, out_meas])

                # --- D. 建立全新的 Excel 檔案寫入資料 ---
                wb = Workbook()
                ws = wb.active
                ws.title = "Processed_Packing"

                # 寫入表頭 (原汁原味)
                for r_idx, r_data in enumerate(header_data, 1):
                    for c_idx, val in enumerate(r_data, 1):
                        if c_idx <= 8: # 保留 A 到 H 欄的寬度
                            ws.cell(row=r_idx, column=c_idx, value=val).font = Font(name='Arial', size=10)

                # 寫入全新的第 12 行標題 (已移除英文品名與 U/M)
                headers_row12 = ["SKU", "Description_of_Goods_(zh)", "Qty", "Net_Weight_(KG)", "Gross_Weight_(KG)", "Measurement_(cm)"]
                for c_idx, val in enumerate(headers_row12, 1):
                    ws.cell(row=12, column=c_idx, value=val).font = Font(name='Arial', size=10, bold=True)

                # 寫入整理好的商品資料
                start_row = 13
                for r_idx, row_data in enumerate(final_output_rows, start_row):
                    for c_idx, val in enumerate(row_data, 1):
                        cell = ws.cell(row=r_idx, column=c_idx, value=val)
                        cell.font = Font(name='Arial', size=10)
                        if c_idx in [3, 4, 5] and val != "":
                            try:
                                cell.value = float(val)
                                cell.number_format = '0.00' if c_idx in [4, 5] else '0'
                            except: pass

                # 寫入最底部的總計
                total_row_idx = start_row + len(final_output_rows)
                qty_display = int(total_qty_sum) if total_qty_sum.is_integer() else round(total_qty_sum, 2)
                ws.cell(row=total_row_idx, column=1, value=f"總箱數:{sku_count}箱").font = Font(name='Arial', size=10, bold=True)
                ws.cell(row=total_row_idx, column=3, value=qty_display).font = Font(name='Arial', size=10, bold=True)
                ws.cell(row=total_row_idx, column=4, value=total_net_sum).font = Font(name='Arial', size=10, bold=True)
                ws.cell(row=total_row_idx, column=5, value=total_gross_sum).font = Font(name='Arial', size=10, bold=True)

                # --- E. 完美復原您要求的合併儲存格 ---
                merge_rules = [
                    ('A1:H1', 'center'), ('A2:H2', 'center'), ('A4:H4', 'center'),
                    ('A3:D3', 'left'), ('E3:H3', 'left'),
                    ('A5:D5', 'left'), ('E5:H5', 'left'),
                    ('A6:H6', 'left'),
                    # 額外補齊下方表頭避免文字被切斷
                    ('A7:D7', 'left'), ('E7:H7', 'left'), 
                    ('A8:D8', 'left'), ('E8:H8', 'left'), 
                    ('A10:D10', 'left'), ('E10:H10', 'left')
                ]
                
                for m_range, align in merge_rules:
                    ws.merge_cells(m_range)
                    top_left = m_range.split(':')[0]
                    ws[top_left].alignment = Alignment(horizontal=align, vertical='center')

                # 設定欄寬與無框線
                col_widths = {'A': 20, 'B': 45, 'C': 8, 'D': 16, 'E': 16, 'F': 18, 'G': 12, 'H': 12}
                for col, width in col_widths.items():
                    ws.column_dimensions[col].width = width

                for row in ws.iter_rows():
                    for cell in row:
                        cell.border = Border() # 徹底清除框線
                        if cell.alignment.vertical is None:
                            cell.alignment = Alignment(vertical='center')

                output = BytesIO()
                wb.save(output)
                st.balloons()
                st.success("✅ 格式重建成功！上方表頭已完美合併置中，資料列完美左移。")
                st.download_button(
                    label="📥 下載排版修正後的 Packing (.xlsx)",
                    data=output.getvalue(),
                    file_name="Global_Packing_Perfect.xlsx",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"系統轉換發生錯誤：{e}")
