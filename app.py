"""批次 Gmail 寄信工具 — Streamlit 網頁介面（本機執行）。

使用流程：
1. 左側先用 Google 帳號登入授權（第一次會跳瀏覽器頁面同意授權）
2. 上傳 Excel 檔案，選出「信箱」欄位
3. 撰寫主旨 / 內文，選擇是否要依 Excel 欄位客製化內容
4. 選擇收件人類型（收件人 / 副本 / 密件副本）與一次寄送人數
5. 選擇立即寄送或排程時間，按下寄送
"""

from __future__ import annotations

from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

import gmail_auth
import mailer

st.set_page_config(page_title="批次 Gmail 寄信工具", page_icon="📧", layout="wide")

RECIPIENT_TYPE_LABELS = {"收件人（一般）": "to", "副本（CC）": "cc", "密件副本（BCC）": "bcc"}


def get_sender_email(service) -> str:
    profile = service.users().getProfile(userId="me").execute()
    return profile["emailAddress"]


# ---------- 側邊欄：Google 帳號授權 ----------
st.sidebar.header("Google 帳號授權")

if "gmail_service" not in st.session_state:
    st.session_state.gmail_service = None
    st.session_state.sender_email = None

if gmail_auth.is_authorized() and st.session_state.gmail_service is None:
    try:
        st.session_state.gmail_service = gmail_auth.get_gmail_service()
        st.session_state.sender_email = get_sender_email(st.session_state.gmail_service)
    except Exception as exc:  # noqa: BLE001
        st.sidebar.error(f"讀取既有授權失敗：{exc}")

if st.session_state.gmail_service is not None:
    st.sidebar.success(f"已登入：{st.session_state.sender_email}")
    if st.sidebar.button("登出（清除授權）"):
        import os

        if os.path.exists(gmail_auth.TOKEN_PATH):
            os.remove(gmail_auth.TOKEN_PATH)
        st.session_state.gmail_service = None
        st.session_state.sender_email = None
        st.rerun()
else:
    st.sidebar.info("尚未登入 Google 帳號")
    if st.sidebar.button("登入 Google 帳號", type="primary"):
        try:
            st.session_state.gmail_service = gmail_auth.get_gmail_service()
            st.session_state.sender_email = get_sender_email(st.session_state.gmail_service)
            st.rerun()
        except FileNotFoundError as exc:
            st.sidebar.error(str(exc))
        except Exception as exc:  # noqa: BLE001
            st.sidebar.error(f"授權失敗：{exc}")

st.title("📧 批次 Gmail 寄信工具")

if st.session_state.gmail_service is None:
    st.warning("請先在左側完成 Google 帳號授權，才能繼續設定寄信內容。")
    st.stop()

# ---------- 上傳 Excel ----------
st.header("1. 上傳 Excel 收件人名單")

with st.expander("📋 Excel 檔案格式說明（第一次上傳前請先看一下）", expanded=False):
    st.markdown(
        """
- 副檔名限定 **.xlsx 或 .xls**（不支援 .csv、.numbers）
- 只會讀取**第一個工作表（分頁）**，如果有多個分頁，其他分頁不會被讀到
- **第一列必須是欄位標題**，資料從第二列開始，欄位名稱、順序不限
- 上傳後可以自己從下拉選單選「哪一欄是信箱」，不需要固定欄名
- 每一格**只能放一個信箱**，一格裡塞多個信箱（例如用 `;` 分隔）不會自動拆開
- 若使用「客製化寄信」，樣板裡的 `{{欄位名稱}}` 要跟 Excel 欄位標題文字**完全一致**（含中英文、空格）
- 信箱欄位是空白的列會被自動略過，不會拿去寄信
        """
    )

uploaded = st.file_uploader("選擇 Excel 檔案 (.xlsx / .xls)", type=["xlsx", "xls"])

if not uploaded:
    st.info("請上傳一個包含收件人信箱的 Excel 檔案。")
    st.stop()

df = pd.read_excel(uploaded)
df = df.dropna(how="all")
st.dataframe(df.head(20), use_container_width=True)
st.caption(f"共讀取到 {len(df)} 筆資料（僅預覽前 20 筆）")

email_col = st.selectbox("哪一欄是收件人信箱？", options=list(df.columns))

# ---------- 撰寫信件 ----------
st.header("2. 撰寫信件內容")

personalize = st.checkbox(
    "客製化寄信（依每筆資料的欄位內容，個別產生信件內容）",
    value=False,
    help="勾選後，主旨/內文中可用 {{欄位名稱}} 帶入該收件人這一列的資料，"
    "系統會針對每個人分別寄一封信。",
)

if personalize:
    cols_hint = "、".join(f"{{{{{c}}}}}" for c in df.columns)
    st.caption(f"可用欄位變數：{cols_hint}")

subject = st.text_input("主旨", placeholder="例如：您好，{{姓名}} 這是本次通知")
body = st.text_area(
    "內文",
    height=220,
    placeholder="例如：\n{{姓名}} 您好，\n\n這是通知內容……",
)

# ---------- 收件人類型與批次 ----------
st.header("3. 收件人類型與寄送批次")

col_a, col_b = st.columns(2)
with col_a:
    recipient_label = st.radio("收件人類型", options=list(RECIPIENT_TYPE_LABELS.keys()))
    recipient_type = RECIPIENT_TYPE_LABELS[recipient_label]

with col_b:
    if personalize:
        st.number_input("一次寄送人數（每人一封信，此設定不適用）", value=1, disabled=True)
        batch_size = 1
    else:
        batch_size = st.number_input(
            "一次寄送人數（同一批收件人共用同一封信）",
            min_value=1,
            max_value=500,
            value=20,
            help="例如設定 20，代表每封信會一次放 20 位收件人在所選欄位（收件人/副本/密件副本）。",
        )

# ---------- 排程 ----------
st.header("4. 寄送時間")

schedule_mode = st.radio("寄送方式", options=["立即寄送", "排程寄送"])

target_datetime = None
if schedule_mode == "排程寄送":
    default_dt = datetime.now() + timedelta(minutes=5)
    d = st.date_input("日期", value=default_dt.date())
    t = st.time_input("時間", value=default_dt.time())
    target_datetime = datetime.combine(d, t)
    if target_datetime <= datetime.now():
        st.error("排程時間必須是未來的時間。")

# ---------- 寄送 ----------
st.header("5. 確認並寄送")

recipients = df.to_dict("records")
valid_recipients = [r for r in recipients if pd.notna(r.get(email_col)) and str(r.get(email_col)).strip()]
st.write(f"有效收件人數：{len(valid_recipients)} / {len(recipients)}")

ready = bool(subject.strip()) and bool(body.strip()) and len(valid_recipients) > 0
if schedule_mode == "排程寄送" and (target_datetime is None or target_datetime <= datetime.now()):
    ready = False

if st.button("🚀 寄送", type="primary", disabled=not ready):
    service = st.session_state.gmail_service
    sender = st.session_state.sender_email

    if schedule_mode == "立即寄送":
        progress = st.progress(0.0)
        status = st.empty()

        def _on_progress(done: int, total: int) -> None:
            progress.progress(done / total if total else 1.0)
            status.text(f"已處理 {done}/{total}")

        report = mailer.send_campaign(
            service=service,
            sender=sender,
            recipients=valid_recipients,
            email_field=email_col,
            subject_template=subject,
            body_template=body,
            personalize=personalize,
            recipient_type=recipient_type,
            batch_size=int(batch_size),
            progress_callback=_on_progress,
        )
        st.success(f"寄送完成：成功 {report.sent_count} 人，失敗 {report.failed_count} 人")
        failed = [r for r in report.results if not r.ok]
        if failed:
            st.error("以下批次寄送失敗：")
            for r in failed:
                st.write(f"- {', '.join(r.recipients)}：{r.error}")
    else:
        def _job():
            mailer.send_campaign(
                service=service,
                sender=sender,
                recipients=valid_recipients,
                email_field=email_col,
                subject_template=subject,
                body_template=body,
                personalize=personalize,
                recipient_type=recipient_type,
                batch_size=int(batch_size),
            )

        mailer.run_at(target_datetime, _job)
        st.success(
            f"已排程於 {target_datetime:%Y-%m-%d %H:%M} 自動寄送，"
            "請保持這個網頁伺服器（終端機視窗）持續執行直到寄送完成。"
        )
