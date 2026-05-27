import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font

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

st.markdown('<p class="big-title">📦 全球 Packing 轉換器</p>', unsafe_allow_html=True)

# 2. 檔案上傳
uploaded_file = st.file_uploader("請上傳客人的 Packing 檔案 (.xlsx)", type=['xlsx'])

if uploaded_file:
    if st.button("🚀 執行智能轉換", use_container_width=True):
        try:
            with st.spinner("正在清理資料與計算重量..."):
                wb = load_workbook(uploaded_file)
                ws = wb.active

                # --- 步驟 0：先殺手級清理，把 SHIPFEE 整列刪除 ---
                # 注意：刪除列必須「由下往上」刪，否則列號(index)會大亂
                for row in range(ws.max_row, 12, -1):
                    desc_val = ws.cell(row=row, column=3).value
                    if desc_val and "SHIPFEE" in str(desc_val).upper():
                        ws.delete_rows(row)

                # --- 步驟 1：預先計算每個 SKU 的總 Qty ---
                sku_totals = {}
                current_sku = None
                total_qty_column_sum = 0 # 最底下的總 Qty (此時已經沒有運費了)
                
                for row in range(13, ws.max_row + 1):
                    sku_val = ws.cell(row=row, column=1).value
                    if sku_val and str(sku_val).strip() != "":
                        current_sku = str(sku_val).strip()
                        if current_sku not in sku_totals:
                            sku_totals[current_sku] = 0
                    
                    qty_val = ws.cell(row=row, column=4).value
                    
                    if qty_val is not None:
                        try:
                            q = float(qty_val)
                            total_qty_column_sum += q # 累加所有有效 Qty
                            
                            if current_sku:
                                sku_totals[current_sku] += q
                        except:
                            pass

                # --- 步驟 2：填入 F, G, H 數據與刪除關鍵字 ---
                total_net_all = 0
                total_gross_all = 0
                sku_count = 0

                for row in range(13, ws.max_row + 1):
                    # 1. 刪除指定關鍵字 (C欄 / column 3)
                    desc_val = ws.cell(row=row, column=3).value
                    if desc_val:
                        desc_str = str(desc_val)
                        for kw in ["【新品】", "★", "【歡樂智多星推薦】"]:
                            desc_str = desc_str.replace(kw, "")
                        ws.cell(row=row, column=3, value=desc_str.strip())

                    # 2. 處理每個 SKU 群組的第一行 (計算重量與材積)
                    sku_val = ws.cell(row=row, column=1).value
                    if sku_val and str(sku_val).strip() != "":
                        current_sku = str(sku_val).strip()
                        sku_count += 1
                        t_qty = sku_totals.get(current_sku, 0)
                        
                        if t_qty > 0:
                            # 毛重=Qty*0.1，淨重=毛重-0.05
                            gross_w = round(t_qty * 0.1, 2)
                            net_w = round(gross_w - 0.05, 2)
                            if net_w < 0: net_w = 0.0

                            # 材積判斷
                            meas = ""
                            if 1 <= t_qty <= 5: meas = "18*19*15"
                            elif 6 <= t_qty <= 10: meas = "23*14*14"
                            elif 11 <= t_qty <= 40: meas = "29*20*20"
                            elif t_qty >= 41: meas = "42*30*25"

                            # 填入 F, G, H 欄 (6, 7, 8)
                            ws.cell(row=row, column=6, value=net_w)
                            ws.cell(row=row, column=7, value=gross_w)
                            ws.cell(row=row, column=8, value=meas)

                            # 累加總重量
                            total_net_all += net_w
                            total_gross_all += gross_w

                # --- 步驟 3：刪除 B 欄與 E 欄，欄位自動左移 ---
                # 必須從右邊先刪除，否則左邊刪除後，右邊的索引會變動
                ws.delete_cols(5) # 刪除原 E 欄 (U/M)
                ws.delete_cols(2) # 刪除原 B 欄 (英文品名)

                # --- 步驟 4：底部計算總和 ---
                max_r = ws.max_row + 1
                # 刪除後，原本的 D, F, G 變成了 C, D, E
                ws.cell(row=max_r, column=1, value=f"總箱數:{sku_count}箱").font = Font(bold=True)
                ws.cell(row=max_r, column=3, value=total_qty_column_sum).font = Font(bold=True)
                ws.cell(row=max_r, column=4, value=round(total_net_all, 2)).font = Font(bold=True)
                ws.cell(row=max_r, column=5, value=round(total_gross_all, 2)).font = Font(bold=True)

                # 存檔供下載
                output = BytesIO()
                wb.save(output)
                st.balloons()
                st.success("✅ Packing 轉換完成！運費項目已徹底移除，格式完美靠左對齊。")
                
                st.download_button(
                    label="📥 下載處理後的 Packing",
                    data=output.getvalue(),
                    file_name="Processed_Packing.xlsx",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"轉換發生錯誤：{e}")
