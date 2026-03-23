"""web_app/config.py - Flaskアプリケーション設定"""
from __future__ import annotations
import pathlib
from datetime import timedelta


APP_VERSION: str = "1.2.0"
APP_RELEASE_DATE: str = "2026-03-22"


class Config:
    """Flask設定クラス"""

    SECRET_KEY: str = 'cc-daily-report-secret-2026'
    DATABASE: str = str(pathlib.Path(__file__).parent.parent / 'db' / 'web_app.db')
    # セッション有効期限: 業務時間相当の8時間
    PERMANENT_SESSION_LIFETIME: timedelta = timedelta(hours=8)
