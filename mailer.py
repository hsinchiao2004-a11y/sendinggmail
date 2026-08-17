"""寄信核心邏輯：組信、客製化內容、收件人類型（收件人/副本/密件副本）、
分批寄送、排程寄送。透過 Gmail 的 SMTP 伺服器 + 應用程式密碼寄信，
不需要 Google Cloud / OAuth 設定。

UI（app.py）只負責蒐集使用者輸入，實際寄信都透過這個模組進行，
方便日後測試或替換介面。
"""

from __future__ import annotations

import re
import smtplib
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from email.mime.text import MIMEText
from email.utils import formataddr
from typing import Callable, Iterable

RecipientType = str  # "to" | "cc" | "bcc"

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465  # SSL

PLACEHOLDER_RE = re.compile(r"\{\{\s*(.+?)\s*\}\}")


def render_template(template: str, row: dict) -> str:
    """把樣板中的 {{欄位名稱}} 換成該筆收件人資料裡對應的值。

    找不到對應欄位時保留原樣（例如 {{公司}}），方便使用者發現打錯欄位名。
    """

    def _replace(match: re.Match) -> str:
        key = match.group(1)
        if key in row and row[key] is not None:
            return str(row[key])
        return match.group(0)

    return PLACEHOLDER_RE.sub(_replace, template)


def test_login(email: str, app_password: str) -> None:
    """驗證帳號密碼是否能登入 Gmail SMTP，登入失敗會拋出例外。"""
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=15) as smtp:
        smtp.login(email, app_password)


def _build_message(
    sender: str,
    sender_name: str | None,
    subject: str,
    body: str,
    to_addrs: list[str],
    cc_addrs: list[str],
) -> MIMEText:
    """組出信件內容（不放 Bcc 表頭，密件副本只在實際收件人清單裡出現）。"""
    message = MIMEText(body, "plain", "utf-8")
    message["from"] = formataddr((sender_name, sender)) if sender_name else sender
    if to_addrs:
        message["to"] = ", ".join(to_addrs)
    if cc_addrs:
        message["cc"] = ", ".join(cc_addrs)
    message["subject"] = subject
    return message


def _dedupe(items: list[str]) -> list[str]:
    """去除重複項目但保留原本順序（避免同一個信箱被重複列在收件清單裡）。"""
    return list(dict.fromkeys(items))


def chunk(items: list, size: int) -> Iterable[list]:
    size = max(1, size)
    for i in range(0, len(items), size):
        yield items[i : i + size]


@dataclass
class SendResult:
    ok: bool
    recipients: list[str]
    error: str | None = None


@dataclass
class SendReport:
    results: list[SendResult] = field(default_factory=list)

    @property
    def sent_count(self) -> int:
        return sum(len(r.recipients) for r in self.results if r.ok)

    @property
    def failed_count(self) -> int:
        return sum(len(r.recipients) for r in self.results if not r.ok)


def send_campaign(
    sender: str,
    app_password: str,
    recipients: list[dict],
    email_field: str,
    subject_template: str,
    body_template: str,
    personalize: bool,
    recipient_type: RecipientType = "to",
    batch_size: int = 20,
    sender_name: str | None = None,
    progress_callback: Callable[[int, int], None] | None = None,
) -> SendReport:
    """寄出整批郵件（透過 Gmail SMTP，一個連線寄完整批）。

    - personalize=True：每人一封信，內容依 {{欄位}} 套用該筆資料，
      batch_size 會被強制視為 1（因為每封信內容都不同）。
    - personalize=False：依 batch_size 把收件人分組，同一組收件人共用同一封信
      （收件人清單放在 recipient_type 指定的欄位：收件人/副本/密件副本）。
    """
    report = SendReport()
    total = len(recipients)
    done = 0

    if personalize:
        groups = [[r] for r in recipients]
    else:
        groups = list(chunk(recipients, batch_size))

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as smtp:
        smtp.login(sender, app_password)

        for group in groups:
            addrs = [str(r[email_field]).strip() for r in group if r.get(email_field)]
            if not addrs:
                done += len(group)
                continue

            if personalize:
                row = group[0]
                subject = render_template(subject_template, row)
                body = render_template(body_template, row)
            else:
                subject = subject_template
                body = body_template

            to_addrs: list[str] = []
            cc_addrs: list[str] = []
            envelope_recipients: list[str] = []
            if recipient_type == "to":
                to_addrs = _dedupe(addrs)
                envelope_recipients = to_addrs
            elif recipient_type == "cc":
                to_addrs = [sender]
                # 如果寄件人自己也在收件清單裡，不重複列在副本欄位，避免同一個人收到兩封
                cc_addrs = _dedupe([a for a in addrs if a != sender])
                envelope_recipients = _dedupe([sender] + addrs)
            elif recipient_type == "bcc":
                to_addrs = [sender]
                envelope_recipients = _dedupe([sender] + addrs)
            else:
                raise ValueError(f"未知的收件人類型：{recipient_type}")

            message = _build_message(sender, sender_name, subject, body, to_addrs, cc_addrs)

            try:
                smtp.sendmail(sender, envelope_recipients, message.as_string())
                report.results.append(SendResult(ok=True, recipients=addrs))
            except smtplib.SMTPException as exc:
                report.results.append(SendResult(ok=False, recipients=addrs, error=str(exc)))

            done += len(group)
            if progress_callback:
                progress_callback(done, total)

    return report


def run_at(target_time: datetime, job: Callable[[], None]) -> threading.Thread:
    """在背景執行緒等到指定時間後執行 job，並回傳該執行緒。

    注意：這個排程只在目前的程式（Streamlit process）持續執行的期間有效，
    如果中途關閉網頁伺服器，排程就會跟著消失。
    """

    def _runner():
        delay = (target_time - datetime.now()).total_seconds()
        if delay > 0:
            time.sleep(delay)
        job()

    thread = threading.Thread(target=_runner, daemon=True)
    thread.start()
    return thread
