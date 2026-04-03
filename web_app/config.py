"""web_app/config.py - Flaskアプリケーション設定"""
from __future__ import annotations
import os
import pathlib
from datetime import timedelta


APP_VERSION: str = "1.8.1"
APP_RELEASE_DATE: str = "2026-04-03"

# プロジェクトルート
_BASE_DIR: pathlib.Path = pathlib.Path(__file__).parent.parent

# SECRET_KEY の永続化ファイル（初回起動時に自動生成）
_SECRET_KEY_FILE: pathlib.Path = _BASE_DIR / "db" / ".secret_key"


def _load_or_generate_secret_key() -> str:
    """SECRET_KEY をファイルから読み込む。なければ生成して保存する。

    Returns:
        str: 64文字の16進ランダム文字列。
    """
    _SECRET_KEY_FILE.parent.mkdir(parents=True, exist_ok=True)
    if _SECRET_KEY_FILE.exists():
        return _SECRET_KEY_FILE.read_text(encoding="utf-8").strip()
    key = os.urandom(32).hex()
    _SECRET_KEY_FILE.write_text(key, encoding="utf-8")
    return key


class Config:
    """Flask設定クラス"""

    SECRET_KEY: str = _load_or_generate_secret_key()
    DATABASE: str = str(_BASE_DIR / "db" / "web_app.db")
    # セッション有効期限: 業務時間相当の8時間
    PERMANENT_SESSION_LIFETIME: timedelta = timedelta(hours=8)
