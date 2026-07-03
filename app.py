import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
import math

# ── 網頁基本設定 ──────────────────────────────────────────────
st.set_page_config(page_title="全球 Packing 轉換器", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #0E1117; }
    .big-title { font-size: 30px !important; font-weight: bold; color: #FFFFFF !important; }
    .stFileUploader section { background-color: #FFFFFF !important; border-radius: 10px; }
    div.stButton > button { background-color: #FFFFFF !important; color: #000000 !important;
        border: 2px solid #000000 !important; height: 50px; font-weight: bold; width: 100%; }
    .stMarkdown p, label { color: #FFFFFF !important; }
    </style>
    """, unsafe_allow_html=True)
st.markdown('<p class="big-title">🌍 全球 Packing 轉換器</p>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("請上傳原始 Packing 檔案（貼上系統packing 頁籤）", type=['xls', 'xlsx'])

# ── 箱子尺寸對照 ──────────────────────────────────────────────
def get_dimensions(pcs):
    try:
        val = int(float(pcs))
        if 1  <= val <= 5:  return "18*19*15"
        elif 6  <= val <= 10: return "23*14*14"
        elif 11 <= val <= 40: return "29*20*20"
        elif val >= 41:       return "42*30*25"
        return ""
    except:
        return ""

# ── 主轉換函式 ────────────────────────────────────────────────
def convert_packing(uploaded_file):
    # 讀取「貼上系統packing」頁籤
    engine = 'xlrd' if uploaded_file.name.lower().endswith('.xls') else 'openpyxl'
    df = pd.read_excel(
        uploaded_file,
        sheet_name='貼上系統packing',
        header=None,
        dtype=str,
        engine=engine
    ).fillna('')

    # --- A. 表頭（row index 0~9，對應 Excel 第1~10行）---
    # 來源欄位：A=col0, D=col3, E=col4, G=col6
    h = {}
    h['r1_A']  = str(df.iloc[0, 0]).strip()
    h['r2_A']  = str(df.iloc[1, 0]).strip()
    h['r3_A']  = str(df.iloc[2, 0]).strip()
    h['r3_E']  = str(df.iloc[2, 4]).strip()
    h['r4_A']  = str(df.iloc[3, 0]).strip()
    h['r5_A']  = str(df.iloc[4, 0]).strip()
    h['r5_E']  = str(df.iloc[4, 4]).strip()
    h['r6_A']  = str(df.iloc[5, 0]).strip()
    h['r7_A']  = str(df.iloc[6, 0]).strip()
    h['r7_E']  = str(df.iloc[6, 4]).strip()   # 客人地址（長文字，自動換行）
    h['r8_A']  = str(df.iloc[7, 0]).strip()   # Shipped by（含換行，自動換行）
    h['r8_E']  = str(df.iloc[7, 4]).strip()
    h['r9_A']  = str(df.iloc[8, 0]).strip()
    h['r9_D']  = str(df.iloc[8, 3]).strip()
    h['r9_G']  = str(df.iloc[8, 6]).strip()
    h['r10_A'] = str(df.iloc[9, 0]).strip()
    h['r10_E'] = str(df.iloc[9, 4]).strip()

    # --- B. 商品資料（row index 12+，col0=SKU, col2=品名zh, col3=Qty）---
    items = []
    for idx in range(12, len(df)):
        col0 = str(df.iloc[idx, 0]).strip()
        if "總箱數" in col0 or "TOTAL" in col0.upper():
            break
        desc_zh = str(df.iloc[idx, 2]).strip()
        if "SHIPFEE" in desc_zh.upper():
            continue
        items.append(df.iloc[idx].tolist())

    # --- C. 群組化 SKU ---
    groups, current = [], []
    for row in items:
        sku = str(row[0]).strip()
        if sku:
            if current:
                groups.append(current)
            current = [row]
        else:
            current.append(row)
    if current:
        groups.append(current)

    final_rows = []
    total_qty = total_net = total_gross = 0.0
    sku_count = len(groups)

    for group in groups:
        group_qty = 0.0
        for r in group:
            try: group_qty += float(r[3])
            except: pass

        total_qty   += group_qty
        gross_w      = round(group_qty * 0.1, 2)
        net_w        = max(0.0, round(gross_w - 0.05, 2))
        total_gross += gross_w
        total_net   += net_w
        meas         = get_dimensions(group_qty)

        for i, r in enumerate(group):
            desc = str(r[2])
            for kw in ["【新品】", "★", "【歡樂智多星推薦】"]:
                desc = desc.replace(kw, "")
            desc = desc.strip()
            final_rows.append([
                r[0] if i == 0 else "",
                desc,
                r[3],
                net_w   if i == 0 else "",
                gross_w if i == 0 else "",
                meas    if i == 0 else "",
            ])

    # --- D. 建立 Excel ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Packing"

    bold12     = Font(name='Calibri', size=12, bold=True)
    thin_s     = Side(style='thin', color='000000')
    thin_border = Border(left=thin_s, right=thin_s, top=thin_s, bottom=thin_s)

    # 欄寬
    for col, w in {'A':18.82,'B':37.09,'C':3.91,'D':15.64,'E':17.18,'F':17.55}.items():
        ws.column_dimensions[col].width = w

    # --- E. 寫入表頭 ---
    def sc(row, col, val, wrap=False, ha='left'):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = bold12
        cell.alignment = Alignment(horizontal=ha, vertical='center', wrap_text=wrap)

    sc(1,  1, h['r1_A'],  ha='center')
    sc(2,  1, h['r2_A'],  ha='center')
    sc(3,  1, h['r3_A'])
    sc(3,  4, h['r3_E'])
    sc(4,  1, h['r4_A'],  ha='center')
    sc(5,  1, h['r5_A'])
    sc(5,  4, h['r5_E'])
    sc(6,  1, h['r6_A'])
    sc(7,  1, h['r7_A'])
    sc(7,  4, h['r7_E'],  wrap=True)   # 客人地址：自動換行
    sc(8,  1, h['r8_A'],  wrap=True)   # Shipped by：自動換行
    sc(8,  4, h['r8_E'])
    sc(9,  1, h['r9_A'])
    sc(9,  3, h['r9_D'])
    sc(9,  6, h['r9_G'])
    sc(10, 1, h['r10_A'])
    sc(10, 4, h['r10_E'])

    # --- F. 合併儲存格 ---
    for rng, ha in [
        ('A1:F1','center'), ('A2:F2','center'), ('A4:F4','center'),
        ('A3:C3','left'),   ('D3:F3','left'),
        ('A5:C5','left'),   ('D5:F5','left'),
        ('A6:F6','left'),
        ('A7:C7','left'),   ('D7:F7','left'),
        ('A8:C8','left'),   ('D8:F8','left'),
        ('A9:B9','left'),   ('C9:E9','left'),
        ('A10:C10','left'), ('D10:F10','left'),
    ]:
        ws.merge_cells(rng)
        tl = rng.split(':')[0]
        existing_wrap = ws[tl].alignment.wrap_text
        ws[tl].alignment = Alignment(horizontal=ha, vertical='center', wrap_text=existing_wrap)

    # --- G. 列高（表頭）---
    for r in range(1, 12):
        ws.row_dimensions[r].height = 14.5

    # Row 7：依地址文字長度，最小 60.3
    r7_lines = max(1, math.ceil(len(h['r7_E'].replace('\n','')) / 22) + h['r7_E'].count('\n'))
    ws.row_dimensions[7].height = max(60.3, r7_lines * 16.5)

    # Row 8：依 Shipped by 文字長度
    r8_lines = max(1, math.ceil(len(h['r8_A'].replace('\n','')) / 28) + h['r8_A'].count('\n'))
    ws.row_dimensions[8].height = max(14.5, r8_lines * 16.5)

    # --- H. 第12行標題 ---
    ws.row_dimensions[12].height = 14.5
    for c, val in enumerate(["SKU","Description_of_Goods_(zh)","Qty","Net_Weight_(KG)","Gross_Weight_(KG)","Measurement_(cm)"], 1):
        cell = ws.cell(row=12, column=c, value=val)
        cell.font = bold12
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='left', vertical='center')

    # --- I. 商品資料列 ---
    for i, row_data in enumerate(final_rows):
        r = 13 + i
        desc = str(row_data[1]) if row_data[1] else ''
        lines = max(1, math.ceil(len(desc) / 18) if desc else 1)
        ws.row_dimensions[r].height = max(14.5, lines * 16.5)

        for c, val in enumerate(row_data, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = bold12
            cell.border = thin_border
            if c == 2:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            elif c <= 3:
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if val != "" and c in [3, 4, 5]:
                try:
                    cell.value = float(val)
                    cell.number_format = '0' if c == 3 else '0.00'
                except:
                    pass

    # --- J. 總計列 ---
    total_row = 13 + len(final_rows)
    ws.row_dimensions[total_row].height = 14.5
    qty_disp = int(total_qty) if total_qty == int(total_qty) else round(total_qty, 2)
    for c, val in enumerate([f"總箱數:{sku_count}箱","", qty_disp, round(total_net,2), round(total_gross,2), ""], 1):
        cell = ws.cell(row=total_row, column=c, value=val)
        cell.font = bold12
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='left' if c<=3 else 'center', vertical='center')

    # 表頭字型補套
    for r in range(1, 12):
        for c in range(1, 7):
            ws.cell(r, c).font = bold12

    # --- K. 輸出為 BytesIO ---
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# ── Streamlit UI ──────────────────────────────────────────────
if uploaded_file:
    if st.button("🚀 執行智能規格轉換", use_container_width=True):
        try:
            with st.spinner("正在計算精確排版..."):
                result = convert_packing(uploaded_file)
            st.balloons()
            st.success("✅ 全球 Packing 轉換成功！")
            st.download_button(
                label="📥 下載精確排版 Packing (.xlsx)",
                data=result,
                file_name="Global_Packing_Perfect.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"系統轉換發生錯誤：{e}")
