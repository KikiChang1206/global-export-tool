import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Alignment

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

st.markdown('<p class="big-title">🌍 全球 Packing 轉換器 (.xls 專用版)</p>', unsafe_allow_html=True)

# 2. 檔案上傳 (精準鎖定 .xls 格式)
uploaded_file = st.file_uploader("請上傳您的原始 Packing 檔案 (.xls)", type=['xls', 'xlsx'])

# 3. 材積自動判別區間邏輯
def get_dimensions(pcs):
    try:
        val = int(float(pcs))
        if 1 <= val <= 5: return "18*19*15"
        elif 6 <= val <= 10: return "23*14*14"
        elif 11 <= val <= 40: return "29*20*20"
        elif val >= 41: return "42*30*25"
        return ""
    except: return ""

# 4. 核心轉換邏輯
if uploaded_file:
    if st.button("🚀 執行智能規格轉換", use_container_width=True):
        try:
            with st.spinner("正在讀取舊版 Excel 並執行欄位重組..."):
                # --- 步驟 A：相容性讀取 ---
                # 針對 .xls 檔案強制指定 xlrd 引擎，將所有內容讀為純文字避免變形
                engine = 'xlrd' if uploaded_file.name.lower().endswith('.xls') else 'openpyxl'
                df = pd.read_excel(uploaded_file, header=None, dtype=str).fillna('').copy()

                # --- 步驟 B：先清空舊的底部總計列 (避免重疊計算) ---
                indices_to_clear = []
                for idx in range(12, len(df)):
                    if "總箱數" in str(df.iloc[idx, 0]) or "TOTAL" in str(df.iloc[idx, 0]).upper():
                        indices_to_clear.append(idx)
                if indices_to_clear:
                    df.drop(index=indices_to_clear, inplace=True)
                    df.reset_index(drop=True, inplace=True)

                # --- 步驟 C：整列徹底刪除 SHIPFEE 運費項目 ---
                rows_to_drop = []
                for idx in range(12, len(df)):
                    desc_zh = str(df.iloc[idx, 2]).strip() # 原始 C 欄 (中文品名)
                    if "SHIPFEE" in desc_zh.upper():
                        rows_to_drop.append(idx)
                
                if rows_to_drop:
                    df.drop(index=rows_to_drop, inplace=True)
                    df.reset_index(drop=True, inplace=True)

                # --- 步驟 D：群組計算重量與材積 (精準對位) ---
                idx = 12  # 第 12 行以下 (Excel第13行) 開始是商品資料
                sku_count = 0
                total_qty_sum = 0.0
                total_net_sum = 0.0
                total_gross_sum = 0.0

                while idx < len(df):
                    sku_val = str(df.iloc[idx, 0]).strip() # 原始 A 欄 (SKU)
                    
                    # 只要 A 欄有值，代表這是一個新箱子群組的起點
                    if sku_val != "":
                        sku_count += 1
                        group_rows = [idx]
                        
                        # 往下尋找同一個群組底下的所有關聯列
                        next_idx = idx + 1
                        while next_idx < len(df):
                            next_sku = str(df.iloc[next_idx, 0]).strip()
                            if next_sku != "": # 遇到下一個 SKU 就切斷
                                break
                            group_rows.append(next_idx)
                            next_idx += 1
                        
                        # 累加這個群組內所有的 Qty (原始 D 欄 / index 3)
                        group_qty = 0.0
                        for r in group_rows:
                            try: group_qty += float(df.iloc[r, 3])
                            except: pass
                        
                        total_qty_sum += group_qty
                        
                        # 依照您提供的正確截圖規格計算：
                        # 毛重(Gross) = Qty * 0.1
                        # 淨重(Net) = 毛重 - 0.05
                        gross_w = round(group_qty * 0.1, 2)
                        net_w = round(gross_w - 0.05, 2)
                        if net_w < 0: net_w = 0.0
                        
                        total_gross_sum += gross_w
                        total_net_sum += net_w
                        meas_str = get_dimensions(group_qty)
                        
                        # 回填至群組的第一行：原始 F 欄(Net), G 欄(Gross), H 欄(Meas)
                        df.iloc[idx, 5] = str(net_w)
                        df.iloc[idx, 6] = str(gross_w)
                        df.iloc[idx, 7] = meas_str
                        
                        # 群組內的其他後續列則保持空白不重複顯示
                        for r in group_rows[1:]:
                            df.iloc[r, 5] = ""
                            df.iloc[r, 6] = ""
                            df.iloc[r, 7] = ""
                            
                        idx = next_idx
                    else:
                        idx += 1

                # --- 步驟 E：全面清理中文品名關鍵字 ---
                for idx in range(12, len(df)):
                    desc_str = str(df.iloc[idx, 2])
                    for kw in ["【新品】", "★", "【歡樂智多星推薦】"]:
                        desc_str = desc_str.replace(kw, "")
                    df.iloc[idx, 2] = desc_str.strip()

                # --- 步驟 F：刪除 B 欄與 E 欄，欄位無縫往前左移 ---
                # 原始：0=SKU, 1=en(刪), 2=zh, 3=Qty, 4=U/M(刪), 5=Net, 6=Gross, 7=Meas
                cols_to_drop = [c for c in [1, 4] if c in df.columns]
                df.drop(columns=cols_to_drop, inplace=True)
                df.columns = range(df.shape[1]) # 重新編排索引，此時：0=SKU, 1=zh, 2=Qty, 3=Net, 4=Gross, 5=Meas

                # --- 步驟 G：在最下方建立全新的總計列 ---
                qty_display = int(total_qty_sum) if total_qty_sum.is_integer() else round(total_qty_sum, 2)
                df.loc[len(df)] = [
                    f"總箱數:{sku_count}箱", 
                    "", 
                    str(qty_display), 
                    f"{total_net_sum:.2f}", 
                    f"{total_gross_sum:.2f}", 
                    ""
                ]

                # --- 步驟 H：利用 openpyxl 建立高質感無框線檔案 ---
                wb = Workbook()
                ws = wb.active
                ws.title = "Processed_Packing"
                
                # 設定精準欄寬
                col_widths = {1: 18, 2: 45, 3: 10, 4: 16, 5: 16, 6: 16}
                for c_idx, width in col_widths.items():
                    ws.column_dimensions[chr(64 + c_idx)].width = width

                for r_idx, row_data in enumerate(df.values, 1):
                    # 動態設定表頭行高與資料行高
                    if r_idx <= 11: ws.row_dimensions[r_idx].height = 18
                    elif r_idx == 12: ws.row_dimensions[r_idx].height = 25
                    else: ws.row_dimensions[r_idx].height = 20

                    for c_idx, val in enumerate(row_data, 1):
                        cell = ws.cell(row=r_idx, column=c_idx)
                        
                        # 資料列的數值轉換與原生格式化 (確保下載後在 Excel 裡是數字格式)
                        if r_idx >= 13 and c_idx in [3, 4, 5]: 
                            try:
                                cell.value = float(val)
                                cell.number_format = '0.00' if c_idx in [4, 5] else '0'
                            except:
                                cell.value = val
                        else:
                            cell.value = val

                        # 樣式配置
                        cell.font = Font(name='Arial', size=10)
                        cell.border = Border() # 完全無框線風格
                        
                        # 對齊設定
                        if r_idx == 12: # 標題列
                            cell.alignment = Alignment(horizontal='left', vertical='center')
                        elif r_idx == len(df): # 總計列
                            cell.font = Font(name='Arial', size=10, bold=True)
                            cell.alignment = Alignment(horizontal='left', vertical='center')
                        else:
                            if c_idx in [3, 4, 5]: # 數據靠右
                                cell.alignment = Alignment(horizontal='right', vertical='center')
                            else:
                                cell.alignment = Alignment(horizontal='left', vertical='center')

                # 輸出
                output = BytesIO()
                wb.save(output)
                st.balloons()
                st.success("✅ 原始 .xls 檔案轉換成功！已完美左移並剔除運費。")
                st.download_button(
                    label="📥 下載處理後的 Packing 檔案 (.xlsx)",
                    data=output.getvalue(),
                    file_name="Global_Packing_Fixed.xlsx",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"系統轉換發生錯誤：{e}")
