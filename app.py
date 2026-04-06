import streamlit as st
import google.generativeai as genai
import openpyxl
import io
import json
import time
import datetime

# 頁面基本設定
st.set_page_config(page_title="丸龜採購單視覺助手", layout="wide")

# --- 1. API Key 安全檢查與模型初始化 ---
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
else:
    st.error("🔑 錯誤：請先在 Streamlit Cloud 的 Settings > Secrets 中設定 GEMINI_API_KEY。")
    st.stop()

st.title("🍣 丸龜採購單自動化助手 (AI 視覺版)")

# --- 🛠️ 除錯工具：檢查可用模型清單 ---
with st.expander("🔍 除錯工具：檢查我的 API Key 可用模型"):
    if st.button("點我獲取模型清單"):
        try:
            available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            st.write("你的 API Key 目前可以使用的模型如下：")
            st.json(available_models)
            st.info("註：如果清單中有 'models/gemini-1.5-flash'，則程式應可正常運作。")
        except Exception as e:
            st.error(f"無法獲取模型清單，可能是 API Key 無效或網路問題：{e}")

st.markdown("---")
st.markdown("""
只要將採購單 PDF 丟上來，AI 就會直接「看」內容並自動填入 Excel。
""")

# 設定樣板路徑 (確保 GitHub 儲存庫裡有 '丸.xlsx')
TEMPLATE_PATH = "丸.xlsx"

# --- 2. 檔案上傳介面 ---
uploaded_pdfs = st.file_uploader("請上傳 PDF 採購單 (可同時選取多份)", type=["pdf"], accept_multiple_files=True)

if st.button("🚀 開始全自動視覺辨識"):
    if uploaded_pdfs:
        try:
            # 讀取 Excel 樣板
            with open(TEMPLATE_PATH, "rb") as f:
                template_bytes = f.read()
            wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
            ws = wb.active
            
            # --- 3. 解析 Excel 結構 ---
            date_cols = {}
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(row=1, column=c).value)
                if "2026-" in val:
                    date_cols[val[:10]] = c
            
            store_rows = {}
            for r in range(1, ws.max_row + 1):
                val = str(ws.cell(row=r, column=2).value).strip()
                if val and val != "None" and val != "店號":
                    store_rows[val] = r
                    store_rows[val.lstrip('0')] = r

            # --- 4. 處理每一份 PDF ---
            # 這裡使用最穩定的模型名稱，若檢查清單顯示不同，請自行修改此處
            model = genai.GenerativeModel('models/gemini-1.5-flash')
            
            progress_bar = st.progress(0)
            for idx, pdf_file in enumerate(uploaded_pdfs):
                st.info(f"正在分析檔案：{pdf_file.name}")
                pdf_bytes = pdf_file.getvalue()
                
                prompt = """
                這是一份採購單。請分析每一頁的內容，並回傳為 JSON 列表。
                1. 『date』: YYYY-MM-DD
                2. 『store』: 三位店號 (如 052)
                3. 『boxes』: 數量欄位的箱數 (純數字)
                格式: [{"page": 1, "date": "2026-04-04", "store": "052", "boxes": 1}]
                """
                
                response = model.generate_content([
                    prompt,
                    {"mime_type": "application/pdf", "data": pdf_bytes}
                ])
                
                try:
                    json_str = response.text.replace("```json", "").replace("```", "").strip()
                    data_list = json.loads(json_str)
                    
                    for item in data_list:
                        d, s, b = item['date'], item['store'], item['boxes']
                        col = date_cols.get(d)
                        row = store_rows.get(s)
                        
                        if col and row:
                            ws.cell(row=row, column=col).value = int(b)
                            st.success(f"✅ 成功：店號 {s} | 日期 {d} | 箱數 {b}")
                        else:
                            st.warning(f"⚠️ 找不到座標：店號 {s}, 日期 {d}")
                except:
                    st.error(f"❌ 辨識失敗：{pdf_file.name}")
                
                progress_bar.progress((idx + 1) / len(uploaded_pdfs))
                time.sleep(1)

            # --- 5. 下載成果 ---
            output = io.BytesIO()
            wb.save(output)
            st.markdown("---")
            st.balloons()
            st.download_button(
                label="📩 下載完成的 Excel 報表",
                data=output.getvalue(),
                file_name=f"丸龜結果_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"發生錯誤：{e}")
    else:
        st.warning("請先上傳 PDF。")
