"""run_production.py - 本番用サーバー起動スクリプト（waitress）

複数ユーザーの同時アクセスに対応した本番用起動スクリプト。
waitress は Windows 環境でも安定動作する WSGI サーバー。

使い方:
    python run_production.py
    python run_production.py --port 8080
    python run_production.py --threads 8
"""
from __future__ import annotations

import argparse
import logging
import sys

from web_app.app import create_app


def parse_args() -> argparse.Namespace:
    """コマンドライン引数を解析して返す。

    Returns:
        argparse.Namespace: 解析済みの引数。
    """
    parser = argparse.ArgumentParser(
        description="予定管理システム 本番サーバー起動"
    )
    parser.add_argument(
        "--host", default="0.0.0.0",
        help="バインドするホスト（デフォルト: 0.0.0.0＝全インターフェース）",
    )
    parser.add_argument(
        "--port", type=int, default=5000,
        help="ポート番号（デフォルト: 5000）",
    )
    parser.add_argument(
        "--threads", type=int, default=4,
        help="ワーカースレッド数（デフォルト: 4、同時接続ユーザー数の目安）",
    )
    return parser.parse_args()


def main() -> None:
    """本番サーバーを起動する。"""
    args = parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    logger = logging.getLogger(__name__)

    try:
        from waitress import serve
    except ImportError:
        logger.error(
            "waitress がインストールされていません。"
            "  pip install waitress で導入してください。"
        )
        sys.exit(1)

    app = create_app()

    logger.info(
        "本番サーバー起動: http://%s:%d  （スレッド数: %d）",
        args.host, args.port, args.threads,
    )
    logger.info("停止するには Ctrl+C を押してください。")

    serve(
        app,
        host=args.host,
        port=args.port,
        threads=args.threads,
        url_scheme="http",
    )


if __name__ == "__main__":
    main()
