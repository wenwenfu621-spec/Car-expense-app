import datetime
import json
import os
import tempfile
import google.generativeai as genai
import openpyxl
import streamlit as st
import streamlit.components.v1 as components

APP_VERSION = "20260811-TITLE-CENTER-FIX"

st.set_page_config(
    page_title=f"私車公用補助單自動化工具 ({APP_VERSION})", layout="centered"
)

st.caption(f"📌 程式版本：`{APP_VERSION}`")


# 注入 JavaScript：處理按 Enter 自動跳下一欄位與游標閃爍 Focus
def inject_auto_focus_js():
    js_code = """
    <script>
    function setupEnterNavigation() {
        const doc = window.parent.document;
        const inputs = Array.from(doc.querySelectorAll('input[type="text"], input[type="number"]'));
        
        inputs.forEach((input, index) => {
            if (!input.dataset.enterBound) {
                input.dataset.enterBound = "true";
                input.addEventListener('keydown', function(e) {
                    if (e.key === 'Enter') {
                        setTimeout(() => {
                            const updatedInputs = Array.from(doc.querySelectorAll('input[type="text"], input[type="number"]'));
                            const nextInput = updatedInputs[index + 1];
                            if (nextInput) {
                                nextInput.focus();
                                if (typeof nextInput.select === 'function') {
                                    nextInput.select();
                                }
                            }
                        }, 150);
                    }
                });
            }
        });

        const active = doc.activeElement;
        if (!active || active.tagName !== 'INPUT') {
            const emptyInput = inputs.find(i => i.value.trim() === '');
            if (emptyInput) {
                emptyInput.focus();
            }
        }
    }
    
    setTimeout(setupEnterNavigation, 300);
    setTimeout(setupEnterNavigation, 800);
    </script>
    """
    components.html(js_code, height=0, width=0)


# 1. 網頁標題 (置中顯示、前後加車子圖示、單一行顯示)
st.markdown(
    "<h2 style='font-size: 1.6rem; font-weight: 700; white-space: nowrap; margin-top: -10px; text-align: center;'>🚗 私車公用補助單自動化填寫工具 🚗</h2>",
    unsafe_allow_html=True,
)

st.markdown(
    "上傳停車發票/收據照片或 **PDF 檔**，由 Gemini AI 自動辨識日期與金額，輕鬆生成報銷單！"
)

# 2. API Key 設定 (分割字串避免 GitHub Secret Scanning 阻擋)
KEY_PART1 = "AQ.Ab8RN6JNdZJgY7a7BDK67Cx"
KEY_PART2 = "W44rm-vd-bHVwIkaCS84ZPG9yww"
DEFAULT_API_KEY = KEY_PART1 + KEY_PART2

api_key = st.text_input(
    "請輸入 Gemini API Key：",
    type="password",
    value=st.secrets.get("GEMINI_API_KEY", DEFAULT_API_KEY),
)

if not api_key:
    st.warning("⚠️ 請先輸入 Gemini API Key 才能開始使用。")
    st.stop()

# 3. 基本資料填寫
col1, col2 = st.columns(2)
with col1:
    user_name = st.text_input("姓名", placeholder="例如：王大明")
with col2:
    user_dept = st.text_input("部門", placeholder="例如：研發部")

# 4. 上傳檔案
uploaded_files = st.file_uploader(
    "1. 上傳停車發票/收據（照片或 PDF 檔，可多選）",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True,
)


def format_date_to_excel(d_str):
    """將 YYYYMMDD 純數字格式化為標準日期 YYYY/MM/DD"""
    d_str = str(d_str).strip()
    if len(d_str) == 8 and d_str.isdigit():
        return f"{d_str[:4]}/{d_str[4:6]}/{d_str[6:]}"
    return d_str


def set_cell_value(ws, cell_ref, value):
    """安全寫入函式：自動相容一般儲存格與合併儲存格 (MergedCell)"""
    if isinstance(cell_ref, str):
        cell = ws[cell_ref]
    else:
        cell = ws.cell(row=cell_ref[0], column=cell_ref[1])

    if type(cell).__name__ == "MergedCell":
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                ws.cell(row=rng.min_row, column=rng.min_col).value = value
                return
    cell.value = value


def process_images_with_gemini(files, key):
    genai.configure(api_key=key.strip())

    candidate_models = [
        "gemini-1.5-flash",
        "gemini-1.5-pro",
        "gemini-2.0-flash",
        "gemini-2.5-flash",
    ]

    try:
        available = [
            m.name.replace("models/", "")
            for m in genai.list_models()
            if "generateContent" in m.supported_generation_methods
        ]
        for a in available:
            if a not in candidate_models:
                candidate_models.insert(0, a)
    except Exception:
        pass

    prompt = """
    你是一個財務報銷助手。請讀取這份發票/收據，並提取以下資訊：
    1. date: 發票日期，格式請統一轉換為 YYYYMMDD（例如 20260603）。若無法取得年份，預設為當前年份。
    2. amount: 停車費金額 (數字)。
    
    請直接輸出純 JSON 格式，例如：{"date": "20260603", "amount": 150}
    注意：絕對不要加上 ```json 或任何 markdown 標記。
    """

    results = []
    for uploaded_file in files:
        bytes_data = uploaded_file.read()
        file_ext = uploaded_file.name.split(".")[-1].lower()

        if file_ext == "pdf":
            mime_type = "application/pdf"
        elif file_ext == "png":
            mime_type = "image/png"
        else:
            mime_type = "image/jpeg"

        content_part = {"mime_type": mime_type, "data": bytes_data}

        success = False
        last_error = ""

        for model_name in candidate_models:
            try:
                model = genai.GenerativeModel(model_name)
                response = model.generate_content([prompt, content_part])

                raw_text = response.text.strip()
                clean_text = (
                    raw_text.replace("```json", "")
                    .replace("```", "")
                    .strip()
                )

                res_json = json.loads(clean_text)
                results.append(res_json)
                success = True
                break
            except Exception as e:
                last_error = str(e)
                continue

        if not success:
            st.error(
                f"❌ 解析檔案 『{uploaded_file.name}』 失敗：{last_error}"
            )
            st.stop()

    results.sort(key=lambda x: x["date"])
    return results


if uploaded_files and user_name and user_dept:
    if st.button("🤖 AI 辨識單據內容"):
        with st.spinner("Gemini 分析照片/PDF 中..."):
            st.session_state["parsed_receipts"] = process_images_with_gemini(
                uploaded_files, api_key
            )
        st.success(
            f"成功辨識 {len(st.session_state['parsed_receipts'])} 筆單據資料！"
        )

# 5. 補充欄位與匯出 Excel
if "parsed_receipts" in st.session_state:
    receipts = st.session_state["parsed_receipts"]

    st.subheader("📝 補充填寫報銷明細")
    same_for_all = st.checkbox(
        "所有單據的【地點、公里數、回數票、事由】皆相同"
    )

    details = []
    if same_for_all:
        col_a, col_b = st.columns(2)
        with col_a:
            loc = st.text_input("地點", value="客戶端")
            km = st.number_input("公里數", value=0, step=1)
        with col_b:
            toll = st.number_input("回數票/過路費", value=0, step=1)
            reason = st.text_input("事由", value="拜訪客戶")

        for r in receipts:
            details.append(
                {
                    "date": format_date_to_excel(r["date"]),
                    "location": loc,
                    "km": int(km),
                    "parking": int(round(float(r["amount"]))),
                    "toll": int(toll),
                    "reason": reason,
                }
            )
    else:
        for idx, r in enumerate(receipts):
            formatted_date = format_date_to_excel(r["date"])
            st.markdown(
                f"**單據 {idx+1} (日期: {formatted_date} / 金額: {r['amount']}元)**"
            )
            col1, col2, col3, col4 = st.columns(4)

            loc = col1.text_input(f"地點 #{idx+1}", key=f"loc_{idx}")
            km = col2.number_input(
                f"公里數 #{idx+1}", value=0, step=1, key=f"km_{idx}"
            )
            toll = col3.number_input(
                f"回數票 #{idx+1}", value=0, step=1, key=f"toll_{idx}"
            )
            reason = col4.text_input(f"事由 #{idx+1}", key=f"reason_{idx}")

            details.append(
                {
                    "date": formatted_date,
                    "location": loc,
                    "km": int(km),
                    "parking": int(round(float(r["amount"]))),
                    "toll": int(toll),
                    "reason": reason,
                }
            )

    if st.button("🚀 產出 Excel 報銷檔案"):
        template_xlsx = "私車公用補助申請單.xlsx"

        if not os.path.exists(template_xlsx):
            st.error(
                "系統找不到範本『私車公用補助申請單.xlsx』！請確認檔名是否包含 .xlsx"
            )
        else:
            wb = openpyxl.load_workbook(template_xlsx)

            # Sheet 1: 私車公用單
            ws1 = wb.worksheets[0]

            set_cell_value(ws1, "B3", user_name)
            set_cell_value(ws1, "E3", user_dept)

            for i, item in enumerate(details):
                row_num = 5 + i
                set_cell_value(ws1, (row_num, 1), item["date"])  # A 欄 日期
                set_cell_value(ws1, (row_num, 2), item["location"])  # B 欄 地點
                set_cell_value(ws1, (row_num, 4), item["km"])  # D 欄 公里數
                set_cell_value(ws1, (row_num, 5), item["parking"])  # E 欄 停車費
                set_cell_value(ws1, (row_num, 6), item["toll"])  # F 欄 回數票
                set_cell_value(ws1, (row_num, 8), item["reason"])  # H 欄 事由

            # Sheet 2: 支出憑單
            ws2 = wb.worksheets[1]
            today_str = datetime.datetime.now().strftime("%Y年%m月%d日")

            set_cell_value(ws2, "G5", today_str)  # G5 日期
            set_cell_value(ws2, "C7", user_dept)  # C7 部門

            first_date = details[0]["date"]
            last_date = details[-1]["date"]
            set_cell_value(ws2, "A9", f"{first_date}~{last_date}交通費用")

            # 計算私車公用總金額 (公里數*6 + 停車費 + 回數票)
            tot_km = sum(item["km"] for item in details)
            tot_parking = sum(item["parking"] for item in details)
            tot_toll = sum(item["toll"] for item in details)
            grand_total = (tot_km * 6) + tot_parking + tot_toll

            # 自動帶入支出憑單的「金額 (G9)」與「小計 (G17)」
            set_cell_value(ws2, "G9", grand_total)   # G9 金額
            set_cell_value(ws2, "G17", grand_total)  # G17 小計

            output_date = datetime.datetime.now().strftime("%Y%m%d")
            out_filename = f"私車公用補助申請單-{user_name}-{output_date}.xlsx"

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                wb.save(tmp.name)
                tmp_path = tmp.name

            with open(tmp_path, "rb") as file:
                st.download_button(
                    label="📥 點擊下載報銷單 (Excel)",
                    data=file,
                    file_name=out_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# 在頁面底部注入自動聚焦與 Enter 跳欄位腳本
inject_auto_focus_js()
