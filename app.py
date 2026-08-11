import datetime
import json
import os
import tempfile
import google.generativeai as genai
import streamlit as st
import xlrd
from xlutils.copy import copy

APP_VERSION = "20260811-ULTIMATE"

st.set_page_config(
    page_title=f"私車公用補助單自動化工具 ({APP_VERSION})", layout="centered"
)

st.caption(f"📌 程式版本：`{APP_VERSION}`")
st.title("🚗 私車公用補助單自動化填寫工具")
st.markdown(
    "上傳停車發票/收據照片或 **PDF 檔**，由 Gemini AI 自動辨識日期與金額，輕鬆生成報銷單！"
)

# 1. API Key 設定
api_key = st.text_input(
    "請輸入 Gemini API Key：",
    type="password",
    value=st.secrets.get("GEMINI_API_KEY", ""),
)

if not api_key:
    st.warning("⚠️ 請先輸入 Gemini API Key 才能開始使用。")
    st.stop()

# 2. 基本資料填寫
col1, col2 = st.columns(2)
with col1:
    user_name = st.text_input("姓名", placeholder="例如：王大明")
with col2:
    user_dept = st.text_input("部門", placeholder="例如：研發部")

# 3. 上傳檔案
uploaded_files = st.file_uploader(
    "1. 上傳停車發票/收據（照片或 PDF 檔，可多選）",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True,
)


def process_images_with_gemini(files, key):
    genai.configure(api_key=key.strip())

    # 候選模型清單 (包含標準名稱與帶前綴名稱)
    candidate_models = [
        "gemini-1.5-flash",
        "gemini-1.5-pro",
        "gemini-2.0-flash",
        "gemini-2.5-flash",
    ]

    # 試圖從帳號動態取得可用模型名稱補補強
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

        # 逐一嘗試候選模型，確保一定能命中一個支援的端點
        for model_name in candidate_models:
            try:
                model = genai.GenerativeModel(model_name)
                response = model.generate_content([prompt, content_part])

                # 清理回傳內容
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
                f"❌ 解析檔案 『{uploaded_file.name}』 失敗，錯誤細節：{last_error}"
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

# 4. 補充欄位與匯出 Excel
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
            km = st.number_input("公里數", value=0.0)
        with col_b:
            toll = st.number_input("回數票/過路費", value=0.0)
            reason = st.text_input("事由", value="拜訪客戶")

        for r in receipts:
            details.append(
                {
                    "date": r["date"],
                    "location": loc,
                    "km": km,
                    "parking": float(r["amount"]),
                    "toll": toll,
                    "reason": reason,
                }
            )
    else:
        for idx, r in enumerate(receipts):
            st.markdown(
                f"**單據 {idx+1} (日期: {r['date']} / 金額: {r['amount']}元)**"
            )
            col1, col2, col3, col4 = st.columns(4)
            loc = col1.text_input(f"地點 #{idx+1}", key=f"loc_{idx}")
            km = col2.number_input(
                f"公里數 #{idx+1}", value=0.0, key=f"km_{idx}"
            )
            toll = col3.number_input(
                f"回數票 #{idx+1}", value=0.0, key=f"toll_{idx}"
            )
            reason = col4.text_input(f"事由 #{idx+1}", key=f"reason_{idx}")

            details.append(
                {
                    "date": r["date"],
                    "location": loc,
                    "km": km,
                    "parking": float(r["amount"]),
                    "toll": toll,
                    "reason": reason,
                }
            )

    if st.button("🚀 產出 Excel 報銷檔案"):
        template_path = "私車公用補助申請單.xls"

        if not os.path.exists(template_path):
            st.error("系統找不到預設範本『私車公用補助申請單.xls』！")
        else:
            rb = xlrd.open_workbook(template_path, formatting_info=True)
            wb = copy(rb)

            # Sheet 1: 私車公用單
            sheet1 = wb.get_sheet(0)
            sheet1.write(2, 1, user_name)
            sheet1.write(2, 3, user_dept)

            total_parking_and_toll = 0
            for i, item in enumerate(details):
                row = 4 + i
                sheet1.write(row, 0, item["date"])
                sheet1.write(row, 1, item["location"])
                sheet1.write(row, 3, item["km"])
                sheet1.write(row, 4, item["parking"])
                sheet1.write(row, 5, item["toll"])
                sheet1.write(row, 7, item["reason"])
                total_parking_and_toll += item["parking"] + item["toll"]

            # Sheet 2: 支出憑單
            sheet2 = wb.get_sheet(1)
            today_str = datetime.datetime.now().strftime("%Y年%m月%d日")
            sheet2.write(4, 6, today_str)
            sheet2.write(6, 2, user_dept)

            first_date = details[0]["date"]
            last_date = details[-1]["date"]
            sheet2.write(8, 0, f"{first_date}~{last_date}交通費用")
            sheet2.write(8, 6, total_parking_and_toll)
            sheet2.write(16, 6, total_parking_and_toll)

            output_date = datetime.datetime.now().strftime("%Y%m%d")
            out_filename = f"私車公用補助申請單-{user_name}-{output_date}.xls"

            with tempfile.NamedTemporaryFile(
                delete=False, suffix=".xls"
            ) as tmp:
                wb.save(tmp.name)
                tmp_path = tmp.name

            with open(tmp_path, "rb") as file:
                st.download_button(
                    label="📥 點擊下載報銷單 (Excel)",
                    data=file,
                    file_name=out_filename,
                    mime="application/vnd.ms-excel",
                )
