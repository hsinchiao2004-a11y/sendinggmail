"""批次 Gmail 寄信工具 — Streamlit 網頁介面（本機執行）。

使用流程：
1. 左側先輸入 Gmail 信箱 + 應用程式密碼登入
2. 上傳 Excel 檔案，選出「信箱」欄位
3. 撰寫主旨 / 內文，選擇是否要依 Excel 欄位客製化內容
4. 選擇收件人類型（收件人 / 副本 / 密件副本）與一次寄送人數
5. 選擇立即寄送或排程時間，按下寄送
"""

from __future__ import annotations

from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

import mailer

st.set_page_config(page_title="批次 Gmail 寄信工具", layout="wide")

RECIPIENT_TYPE_LABELS = {"收件人（一般）": "to", "副本（CC）": "cc", "密件副本（BCC）": "bcc"}


def excel_col_letter(index: int) -> str:
    """把 0 起始的欄位索引轉成 Excel 樣式的欄位字母（0→A、25→Z、26→AA……）。"""
    index += 1
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters

# ---------- 全域樣式 ----------
CSS = """
<style>
:root {
    --ink: #1a1a1e;
    --muted: #6b6b76;
    --line: #e8e8ec;
    --card: #ffffff;
    --accent: #4f46e5;
}

html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter",
        "PingFang TC", "Microsoft JhengHei", sans-serif;
}

.block-container {
    max-width: 900px;
    padding-top: 2.5rem;
    padding-bottom: 5rem;
}

h1 {
    font-weight: 700;
    letter-spacing: -0.02em;
    margin-bottom: 0.1rem !important;
    text-align: center;
}

.hero-tags {
    display: flex;
    justify-content: center;
    flex-wrap: wrap;
    gap: 0.5rem;
    margin-bottom: 2.2rem;
}

.hero-tag {
    background: #f6f6f8;
    color: var(--muted);
    font-size: 0.84rem;
    font-weight: 500;
    padding: 0.32rem 0.85rem;
    border-radius: 999px;
    border: 1px solid var(--line);
}

.var-box {
    background: #f6f6ff;
    border: 1px solid #e0e0fb;
    border-radius: 10px;
    padding: 0.9rem 1rem;
    margin: 0.6rem 0 1rem 0;
}
.var-box p {
    margin: 0 0 0.5rem 0;
    font-size: 0.86rem;
    color: var(--ink);
}
.var-chips {
    display: flex;
    flex-wrap: wrap;
    gap: 0.4rem;
    margin-bottom: 0.6rem;
}
.var-chip {
    background: white;
    border: 1px solid var(--accent);
    color: var(--accent);
    font-family: ui-monospace, "SFMono-Regular", Menlo, monospace;
    font-size: 0.82rem;
    font-weight: 600;
    padding: 0.15rem 0.6rem;
    border-radius: 6px;
}
.var-example {
    font-size: 0.82rem;
    color: var(--muted);
    margin: 0;
}

.step-header {
    display: flex;
    align-items: center;
    justify-content: flex-start;
    gap: 0.65rem;
    margin: 2.2rem 0 0.7rem 0;
    text-align: left;
}

.step-badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 26px;
    height: 26px;
    border-radius: 8px;
    background: var(--accent);
    color: white;
    font-size: 0.82rem;
    font-weight: 600;
    flex-shrink: 0;
}

.step-title {
    font-size: 1.12rem;
    font-weight: 600;
    color: var(--ink);
    letter-spacing: -0.01em;
    text-align: left;
}

.step-subtitle {
    font-size: 0.86rem;
    color: var(--muted);
    margin-top: 0.1rem;
    text-align: left;
}

.info-list ul {
    margin: 0;
    padding-left: 1.2rem;
}
.info-list li {
    color: var(--ink);
    font-size: 0.88rem;
    line-height: 1.65;
}
.info-list code {
    color: var(--ink);
    background: #f6f6f8;
    border-radius: 4px;
    padding: 0.05rem 0.35rem;
    font-size: 0.83rem;
}

/* 卡片容器（st.container(border=True) 產生的區塊） */
div[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 14px !important;
    border: 1px solid var(--line) !important;
    background: var(--card);
    box-shadow: 0 1px 2px rgba(20, 20, 30, 0.03);
}
div[data-testid="stVerticalBlockBorderWrapper"] > div {
    padding: 1.4rem 1.5rem;
}

/* 按鈕 */
.stButton > button, .stFormSubmitButton > button {
    border-radius: 9px;
    font-weight: 600;
    border: none;
}
.stButton > button[kind="primary"], .stFormSubmitButton > button[kind="primary"] {
    background: var(--accent);
}
.stButton > button[kind="primary"]:hover, .stFormSubmitButton > button[kind="primary"]:hover {
    background: #4338ca;
}

/* 輸入框 */
.stTextInput input, .stTextArea textarea, .stNumberInput input {
    border-radius: 8px !important;
}

/* 進度條 */
.stProgress > div > div > div {
    background-color: var(--accent) !important;
}

section[data-testid="stSidebar"] {
    border-right: 1px solid var(--line);
}

[data-testid="stExpander"] {
    border-radius: 10px !important;
    border: 1px solid var(--line) !important;
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def step_header(badge: str, title: str, subtitle: str = "") -> None:
    subtitle_html = f'<div class="step-subtitle">{subtitle}</div>' if subtitle else ""
    html = (
        f'<div class="step-header"><span class="step-badge">{badge}</span>'
        f'<div><div class="step-title">{title}</div>{subtitle_html}</div></div>'
    )
    st.markdown(html, unsafe_allow_html=True)


if "sender_email" not in st.session_state:
    st.session_state.sender_email = None
    st.session_state.app_password = None
    st.session_state.sender_name = None

st.title("批次 Gmail 寄信工具")
st.markdown(
    '<div class="hero-tags">'
    '<span class="hero-tag">上傳 Excel 名單</span>'
    '<span class="hero-tag">客製化內容</span>'
    '<span class="hero-tag">設定收件人類型與寄送時間</span>'
    "</div>",
    unsafe_allow_html=True,
)

# ---------- Gmail 帳號登入（置中顯示，登入前一直看得到） ----------
if not st.session_state.sender_email:
    left, mid, right = st.columns([1, 2, 1])
    with mid:
        with st.container(border=True):
            st.markdown('<div class="step-title">登入 Gmail 帳號</div>', unsafe_allow_html=True)

            name_input = st.text_input(
                "顯示名稱",
                placeholder="例如：陳小美",
                help="收件人看到的寄件人名稱，不填的話對方只會看到你的 email 地址。",
            )
            email_input = st.text_input("Gmail 信箱", placeholder="you@gmail.com")
            password_input = st.text_input("應用程式密碼（16 碼）", type="password")
            submitted = st.button("登入", type="primary", use_container_width=True)

            if submitted:
                if not email_input.strip() or not password_input.strip():
                    st.error("信箱和應用程式密碼都要填。")
                else:
                    try:
                        mailer.test_login(email_input.strip(), password_input.strip())
                        st.session_state.sender_email = email_input.strip()
                        st.session_state.app_password = password_input.strip()
                        st.session_state.sender_name = name_input.strip()
                        st.rerun()
                    except Exception as exc:  # noqa: BLE001
                        st.error(f"登入失敗，請確認信箱與應用程式密碼是否正確：{exc}")

            st.divider()
            st.markdown(
                "**還沒申請應用程式密碼？**\n\n"
                "1. 到 Google 帳號設定開啟「兩步驟驗證」\n"
                "2. 搜尋「應用程式密碼」，建立一組新的（名稱隨便打）\n"
                "3. 複製產生的 16 碼密碼貼到上面（不是你平常登入的密碼）\n\n"
                "詳細步驟請看 README。"
            )
    st.stop()

status_col, logout_col = st.columns([5, 1])
with status_col:
    st.success(f"已登入：{st.session_state.sender_email}")
with logout_col:
    if st.button("登出", use_container_width=True):
        st.session_state.sender_email = None
        st.session_state.app_password = None
        st.session_state.sender_name = None
        st.rerun()

# ---------- 上傳 Excel ----------
step_header("1", "上傳 Excel 收件人名單")

with st.container(border=True):
    with st.expander("Excel 檔案格式說明（第一次上傳前請先看一下）", expanded=False):
        st.markdown(
            """
<div class="info-list">
<ul>
<li>副檔名限定 <code>.xlsx</code> 或 <code>.xls</code>（不支援 <code>.csv</code>、<code>.numbers</code>）</li>
<li>只會讀取第一個工作表（分頁），如果有多個分頁，其他分頁不會被讀到</li>
<li>第一列必須是欄位標題，資料從第二列開始，欄位名稱、順序不限</li>
<li>上傳後可以自己從下拉選單選「哪一欄是信箱」，不需要固定欄名</li>
<li>每一格只能放一個信箱，一格裡塞多個信箱（例如用 <code>;</code> 分隔）不會自動拆開</li>
<li>若使用「客製化寄信」，樣板裡的 <code>{{欄位名稱}}</code> 要跟 Excel 欄位標題文字完全一致（含中英文、空格）</li>
<li>信箱欄位是空白的列會被自動略過，不會拿去寄信</li>
</ul>
</div>
            """,
            unsafe_allow_html=True,
        )

    uploaded = st.file_uploader("選擇 Excel 檔案 (.xlsx / .xls)", type=["xlsx", "xls"])

    if not uploaded:
        st.info("請上傳一個包含收件人信箱的 Excel 檔案。")
        st.stop()

    df = pd.read_excel(uploaded)
    df = df.dropna(how="all")
    # 欄位標題一律當成文字處理（例如標題只是數字「7」時，pandas 會讀成整數，
    # 導致 {{7}} 這種文字變數永遠比對不到，客製化就會失效）
    df.columns = [str(c) for c in df.columns]

    col_letters = {col: excel_col_letter(i) for i, col in enumerate(df.columns)}

    df_preview = df.head(20).copy()
    df_preview.columns = [f"{col_letters[c]}（{c}）" for c in df.columns]
    st.dataframe(df_preview, use_container_width=True)
    st.caption(f"共讀取到 {len(df)} 筆資料（僅預覽前 20 筆；欄位前的字母對應原始 Excel 的欄位順序 A、B、C……）")

    email_col = st.selectbox(
        "哪一欄是收件人信箱？",
        options=list(df.columns),
        format_func=lambda c: f"{col_letters[c]}（{c}）",
    )

# ---------- 撰寫信件 ----------
step_header("2", "撰寫信件內容")

with st.container(border=True):
    personalize = st.checkbox(
        "客製化寄信（依每筆資料的欄位內容，個別產生信件內容）",
        value=False,
        help="勾選後，主旨/內文中可用 {{欄位名稱}} 帶入該收件人這一列的資料，"
        "系統會針對每個人分別寄一封信。",
    )

    if personalize:
        chips_html = "".join(f'<span class="var-chip">{{{{{c}}}}}</span>' for c in df.columns)
        example_col = df.columns[0]
        st.markdown(
            '<div class="var-box">'
            "<p>把下面這些文字直接複製、貼到下面「主旨」或「內文」欄位裡，"
            "系統寄信時會自動換成該收件人那一列的實際資料：</p>"
            f'<div class="var-chips">{chips_html}</div>'
            f'<p class="var-example">範例：如果內文打「{{{{{example_col}}}}} 您好」，'
            f"某一列的「{example_col}」欄位值是「王小明」，這位收件人收到的內文就會變成「王小明 您好」。</p>"
            "</div>",
            unsafe_allow_html=True,
        )

    subject = st.text_input("主旨", placeholder="例如：您好，{{姓名}} 這是本次通知")
    body = st.text_area(
        "內文",
        height=220,
        placeholder="例如：\n{{姓名}} 您好，\n\n這是通知內容……",
    )

# ---------- 收件人類型與批次 ----------
step_header("3", "收件人類型與寄送批次")

with st.container(border=True):
    col_a, col_b = st.columns(2)
    with col_a:
        if personalize:
            st.radio(
                "收件人類型",
                options=list(RECIPIENT_TYPE_LABELS.keys()),
                index=0,
                disabled=True,
                help="已勾選「客製化寄信」：每封信本來就只寄給一個人，"
                "不需要（也不應該）再用副本/密件副本隱藏，固定用「收件人（一般）」直接寄給本人。",
            )
            st.caption("已勾選客製化寄信，固定用「收件人（一般）」，此欄位鎖住不能改")
            recipient_type = "to"
        else:
            recipient_label = st.radio("收件人類型", options=list(RECIPIENT_TYPE_LABELS.keys()))
            recipient_type = RECIPIENT_TYPE_LABELS[recipient_label]

    with col_b:
        if personalize:
            st.number_input(
                "一次寄送人數",
                value=1,
                disabled=True,
                help="已勾選「客製化寄信」：每個人收到的內容都不一樣，"
                "所以無法把多人合併在同一封信裡，固定為每人各寄一封。",
            )
            st.caption("已勾選客製化寄信，固定每人各寄一封，此欄位鎖住不能改")
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
step_header("4", "寄送時間")

with st.container(border=True):
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
step_header("5", "確認並寄送")

with st.container(border=True):
    recipients = df.to_dict("records")
    non_empty = [r for r in recipients if pd.notna(r.get(email_col)) and str(r.get(email_col)).strip()]

    seen_emails: set[str] = set()
    valid_recipients = []
    for r in non_empty:
        addr = str(r[email_col]).strip().lower()
        if addr in seen_emails:
            continue
        seen_emails.add(addr)
        valid_recipients.append(r)

    duplicate_count = len(non_empty) - len(valid_recipients)
    st.write(f"有效收件人數：{len(valid_recipients)} / {len(recipients)}")
    if duplicate_count:
        st.caption(f"偵測到 {duplicate_count} 筆信箱與前面重複，已自動略過，避免同一個人收到多封。")

    ready = bool(subject.strip()) and bool(body.strip()) and len(valid_recipients) > 0
    if schedule_mode == "排程寄送" and (target_datetime is None or target_datetime <= datetime.now()):
        ready = False

    if st.button("寄送", type="primary", disabled=not ready):
        sender = st.session_state.sender_email
        app_password = st.session_state.app_password
        sender_name = st.session_state.sender_name

        if schedule_mode == "立即寄送":
            progress = st.progress(0.0)
            status = st.empty()

            def _on_progress(done: int, total: int) -> None:
                progress.progress(done / total if total else 1.0)
                status.text(f"已處理 {done}/{total}")

            report = mailer.send_campaign(
                sender=sender,
                app_password=app_password,
                recipients=valid_recipients,
                email_field=email_col,
                subject_template=subject,
                body_template=body,
                personalize=personalize,
                recipient_type=recipient_type,
                batch_size=int(batch_size),
                sender_name=sender_name,
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
                    sender=sender,
                    app_password=app_password,
                    recipients=valid_recipients,
                    email_field=email_col,
                    subject_template=subject,
                    body_template=body,
                    personalize=personalize,
                    recipient_type=recipient_type,
                    batch_size=int(batch_size),
                    sender_name=sender_name,
                )

            mailer.run_at(target_datetime, _job)
            st.success(
                f"已排程於 {target_datetime:%Y-%m-%d %H:%M} 自動寄送，"
                "請保持這個網頁伺服器（終端機視窗）持續執行直到寄送完成。"
            )
