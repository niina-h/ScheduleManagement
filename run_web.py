"""run_web.py - 予定管理システム 開発用サーバー起動スクリプト"""
from __future__ import annotations

import logging

from web_app.app import create_app

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
# 起動時の冗長なログを抑制
logging.getLogger("werkzeug").setLevel(logging.ERROR)

import flask.cli
flask.cli.show_server_banner = lambda *args, **kwargs: None

app = create_app()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
