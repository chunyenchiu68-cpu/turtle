import streamlit as st
import google.generativeai as genai
import openpyxl
import io
import json

# --- 關鍵修正區：確保讀取 Secrets ---
try:
    # 優先讀取 Streamlit Secrets
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
except Exception as e:
    st.error(f"❌ 找不到 API Key，請在 Secrets 設定。錯誤: {e}")
    st.stop() # 停止執行後續程式

model = genai.GenerativeModel('gemini-1.5-flash')
# --- 後面接原本的程式碼 ---

st.title("🍣 丸龜採購單自動化助手")
st.write("利用 AI 視覺辨識技術，自動填寫箱數，不再受亂碼困擾。")

# 檔案上傳區
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. 上傳 Excel 樣板 (丸.xlsx)", type=["xlsx"])
with col2:
    uploaded_pdfs = st.file_uploader("2. 上傳 PDF 採購單 (可多選)", type=["pdf"], accept_multiple_files=True)

if st.button("🚀 開始全自動辨識與填寫"):
    if uploaded_excel and uploaded_pdfs:
        # 讀取 Excel
        wb = openpyxl.load_workbook(io.BytesIO(uploaded_excel.read()))
        ws = wb.active
        
        # 建立日期與店號定位地圖 (基於前 5 列日期與 B 欄店號)
        date_cols = {}
        for r in range(1, 6):
            for c in range(1, ws.max_column + 1):
                v = str(ws.cell(row=r, column=c).value)
                if "2026-" in v: date_cols[v[:10]] = c
        
        store_rows = {}
        for r in range(1, ws.max_row + 1):
            v = str(ws.cell(row=r, column=2).value).strip()
            if v and v != "None": store_rows[v] = r

        # 處理 PDF
        for pdf_file in uploaded_pdfs:
            st.info(f"正在辨識: {pdf_file.name}...")
            pdf_data = pdf_file.getvalue()
            
            # 叫 Gemini 視覺辨識每一頁
            prompt = """這是一份採購單。請依序分析每一頁，輸出為 JSON 格式列表。
            每個物件包含: "page": 頁碼, "date": "YYYY-MM-DD", "store": "三位店號", "boxes": 箱數(純數字)。
            範例: [{"page": 1, "date": "2026-04-04", "store": "052", "boxes": 1}]"""
            
            response = model.generate_content([prompt, {"mime_type": "application/pdf", "data": pdf_data}])
            
            try:
                # 去除 Markdown 標籤並解析 JSON
                clean_json = response.text.replace("```json", "").replace("```", "").strip()
                results = json.loads(clean_json)
                
                for item in results:
                    d, s, b = item['date'], item['store'], item['boxes']
                    # 執行填寫
                    target_col = date_cols.get(d)
                    target_row = store_row_map_get_logic = store_rows.get(s) or store_rows.get(s.lstrip('0'))
                    
                    if target_col and target_row:
                        ws.cell(row=target_row, column=target_col).value = int(b)
                        st.success(f"✅ {pdf_file.name} 第 {item['page']} 頁: 店號 {s} / {d} / {b} 箱 填寫成功")
                    else:
                        st.warning(f"⚠️ 找不到座標: 店號 {s} 日期 {d}")
            except:
                st.error(f"❌ {pdf_file.name} 辨識格式有誤，AI 回應: {response.text[:100]}")

        # 下載成果
        output = io.BytesIO()
        wb.save(output)
        st.download_button("📩 下載填寫完成的檔案", output.getvalue(), "丸龜自動化結果.xlsx")
    else:
        st.warning("請先上傳 Excel 與 PDF 檔案。")
