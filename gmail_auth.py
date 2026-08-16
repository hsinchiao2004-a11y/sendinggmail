"""Gmail API OAuth 授權模組。

第一次執行時會開啟瀏覽器讓使用者登入 Google 帳號並同意授權，
之後會把授權結果（refresh token）存到 token.json，下次執行時
直接讀取、不需要重新登入（除非權限被撤銷或 token 過期失效）。

使用前置作業：
1. 到 https://console.cloud.google.com/ 建立專案並啟用 Gmail API
2. 建立 OAuth 用戶端 ID（應用程式類型選「電腦版應用程式 / Desktop app」）
3. 下載憑證 JSON，改名為 credentials.json，放在專案根目錄
"""

from __future__ import annotations

import os

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# 只要求「寄信」權限，不讀取、不刪除信件
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

CREDENTIALS_PATH = "credentials.json"
TOKEN_PATH = "token.json"


def get_credentials() -> Credentials:
    """取得（必要時刷新或重新走一次 OAuth 流程）使用者授權憑證。"""
    creds: Credentials | None = None

    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_PATH):
                raise FileNotFoundError(
                    f"找不到 {CREDENTIALS_PATH}，請先到 Google Cloud Console "
                    "建立 OAuth 憑證並下載後放在專案根目錄（詳見 README）。"
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_PATH, "w", encoding="utf-8") as token_file:
            token_file.write(creds.to_json())

    return creds


def get_gmail_service():
    """回傳已授權、可直接呼叫的 Gmail API service 物件。"""
    creds = get_credentials()
    return build("gmail", "v1", credentials=creds)


def is_authorized() -> bool:
    """檢查目前是否已經有有效（或可刷新）的授權，不會觸發互動式登入。"""
    if not os.path.exists(TOKEN_PATH):
        return False
    try:
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    except Exception:
        return False
    return bool(creds and (creds.valid or (creds.expired and creds.refresh_token)))
