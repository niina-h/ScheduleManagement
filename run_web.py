"""run_web.py - CC日報管理アプリ Webサーバー起動スクリプト"""
from __future__ import annotations

import logging

from web_app.app import create_app

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)

app = create_app()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
