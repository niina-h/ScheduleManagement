# CLAUDE.md — Claude 行動規約

本ファイルは Claude が最優先で従う行動規約である。

> **準拠元**：コーディング規約 令和6年4月08日 商品開発室

---

## 1. 絶対ルール

### 1.1 完全日本語対応

- 応答・コメント・コミットメッセージ・ドキュメントはすべて日本語で行う
- 英語表記が許可されるのは **変数名・関数名・クラス名などの識別子のみ**

### 1.2 確認ポリシー

**確認不要**：コード読み取り・テスト実行・`git status/diff/log`・ファイル読み取り（`.env`除く）
**確認必須**：ファイル削除/上書き/移動・`.env`アクセス・DB変更・外部サービスへのデータ送信

---

## 2. 命名規則（Python）

| 種別 | 規則 | 例 |
|------|------|-----|
| 変数名・関数名 | スネークケース | `user_name`, `get_user_data()` |
| クラス名 | パスカルケース | `UserInfo` |
| 定数 | 大文字スネークケース | `MAX_RETRY_COUNT` |
| グローバル変数 | 接頭辞 `g` | `gVariableName` |
| モジュール変数 | 接頭辞 `m` | `mVariableName` |
| UI要素 | 種別接頭辞 | `txtUserName`, `btnSubmit`, `lblStatus`, `chkAgree`, `rdoMale`, `cmbCountry`, `lstItems`, `grpLogin`, `pnlMain`, `imgLogo`, `tabMain` |
| コレクション | 接頭辞 `col` | `colVariableName` |
| 機能（feature）フォルダ | スネークケース・動作+対象 | `search_tool`, `drawing_a` |

---

## 3. コード品質ルール

- **インデント**：スペース4つ（PEP8準拠）。タブ禁止
- **1行1命令**。関数定義・制御構文の末尾は `:` を付与
- **関数は1画面以内**。ネストは **最大3階層**
- **ラムダ式の多用禁止**。明示的な関数定義を推奨
- **コメント**：処理の直前行に日本語で記述。半角カタカナ禁止
- **リソース管理**：ファイル・ネットワーク等の外部リソースは `with` 文で管理

---

## 4. クラス設計

- **1クラス = 1責務**（単一責任の原則）
- **メンバ構成順序**：定数/変数 → コンストラクタ → Public メソッド → Private メソッド
- **ファイルヘッダー必須**：全ファイル先頭に `# ファイル名：` と `# 概要：` を記載
- **docstring 必須**：すべての公開メソッドに記載

---

## 5. フォルダ配置ルール（マルチ機能対応）

```
project_root/
├─ main.py  CLAUDE.md  README.md  requirements.txt  .gitignore
├─ src/
│  ├─ common/util/file_util.py   # 全機能共有（model/, repository/ 等を追加可能）
│  └─ features/
│     ├─ search_tool/            # 1ツール = 1サブパッケージ
│     │   controller.py          # CLI 起動 run(args)
│     │   service.py             # 業務ロジック
│     │   model.py               # ツール固有モデル
│     │   api.py                 # Web API ルーター
│     └─ web/                    # FastAPI サーバ（認証＋各 feature 集約）
├─ tests/                        # pytest（src/ と同じ階層で配置）
│  ├─ common/                    # common 配下のテスト
│  └─ features/<tool>/           # 各 feature のテスト
├─ input/  output/               # 入出力（中身は Git 除外）
├─ docs/                         # ドキュメント
├─ resources/
│  ├─ config/  image/            # 設定・画像置き場
│  └─ web_template/              # React フロント（不要なら削除可）
│     └─ src/features/search_tool/  # Python 側と対称
├─ build/                        # ビルド成果物
└─ .claude/                      # Claude Code 設定
```

### 原則

- **1ツール = 1サブパッケージ**：`src/features/<tool>/` 配下にまとめる
- **依存方向は features → common の一方向**。features 間の相互参照は禁止
- **新ツール追加**：`main.py` の `_FEATURES` に1行 + 必要なら `web/app.py` で `include_router` を1行
- **Python と Web は対で配置**：`src/features/<tool>/api.py` ⇔ `resources/web_template/src/features/<tool>/`
- **ルート直下に置くのは** `main.py` / `CLAUDE.md` / `README.md` / `requirements.txt` / `.gitignore` のみ
- 入力 → `input/`、出力 → `output/`

---

## 6. セキュリティ（情報流出防止）

- `.env`, `*.pem`, `*.key`, `secrets.json` は読み書きとも禁止（`.claude/settings.json` の deny で遮断）
- `curl`, `wget`, `scp`, `ssh`, `rsync` 等の外部送信は禁止。データを外部へ出さない
- `git push`・`pip/npm install`・`WebSearch` は実行前に確認（ask）
- `rm -rf`, `mkfs`, `dd`, `sudo` 等の破壊的・権限昇格コマンドは禁止
- ユーザー入力はサニタイズ必須（SQLインジェクション・OSコマンド注入対策）
- 外部から受け取るパスは `FileUtil.ensure_within()` でパストラバーサル対策
- ログにパスワード・APIキー・PIIを含めない
- 機密案件は `/sandbox`（ファイル・ネットワーク分離）の併用を推奨

---

## 7. エラーハンドリング・ログ・テスト

- `try/except` で例外種別に応じた処理。例外の握りつぶし禁止
- `logging` モジュールでログレベル（`INFO`, `WARNING`, `ERROR`）を使い分ける
- Excel操作は `pywin32` ライブラリ経由のみ。SQL文はストアドプロシージャで管理
- テストは `pytest` で `tests/` フォルダに配置。`tests/features/<tool_name>/test_*.py` の階層を推奨

---

## 8. 詳細ルール参照

本ファイルの簡潔さを保つため、詳細は以下を参照：

- **PR前チェック** → `/pr-check` スキルを実行
- **コードレビュー** → `/code-review` スキルを実行
- **新規ファイル作成** → `/new-file` スキルを実行
- **リファクタリング** → `/refactor` スキルを実行
- **各ディレクトリ固有ルール** → 各フォルダ内の `CLAUDE.md` を参照

---

以上を **CLAUDE.md（Claude 行動規約）** とする。
