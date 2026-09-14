# Claude Code 向け指示書 — SNS管理ツール リリース作業

本ドキュメントは、SNS管理ツールを **予定管理システム(ScheduleManagement)と同じ運用端末** にリリースする際、Claude Code に対して与える標準指示書である。
既存の予定管理システム(RELEASE_GUIDE.md / UPDATE_GUIDE.md)を参考に、SNS管理ツール向けに再構成している。

---

## 0. 前提条件（Claude Code に必ず伝えること）

| 項目 | 内容 |
|------|------|
| プロジェクト名 | SNS管理ツール (SnsManagementTool 想定) |
| バージョン管理方式 | **Git 未使用**。プロジェクトフォルダ丸ごとバックアップで管理 |
| 対象運用端末 | 予定管理システム稼働中の同一端末 (Windows 10/11) |
| 開発元パス例 | `C:\DEV\SnsManagementTool` |
| 本番配置先例 | `C:\Apps\SnsManagementTool` |
| バックアップ保管先例 | `D:\Backup\SnsManagementTool\` (別ドライブ推奨) |
| 使用ポート | **予定管理システム(5000)と重複しない番号を使用**（例: 5010） |

---

## 1. Claude Code への指示テンプレート

以下をそのままコピーしてプロンプトとして使用する。

```
# 依頼内容
SNS管理ツールを運用端末にリリースしてください。
以下の手順に沿って、確認を挟みながら段階的に進めてください。

## 参考資料
- ScheduleManagement の deploy/RELEASE_GUIDE.md
- ScheduleManagement の deploy/UPDATE_GUIDE.md
- 本指示書 (01_ClaudeCode向け_SNS管理ツール_リリース指示書.md)

## 実行環境情報
- 開発元パス: C:\DEV\SnsManagementTool
- 本番配置先: C:\Apps\SnsManagementTool
- バックアップ先: D:\Backup\SnsManagementTool\
- 使用ポート: 5010（予定管理システム5000と重複回避）
- タスクスケジューラ登録名: SnsManagementServer

## 必ず守ること
- CLAUDE.md の行動規約（日本語対応・確認ポリシー・セキュリティ）を遵守
- .env / .secret_key / DB ファイルは読み書き禁止
- 予定管理システム(ポート5000)の稼働を絶対に止めない
- 破壊的操作（削除・上書き）は必ず事前確認
- 各フェーズ終了後に「次に進んでよいか」ユーザーに確認
```

---

## 2. Claude Code に実行させる作業フロー（フェーズ別）

### フェーズ A — 事前調査（読み取り専用）

Claude Code は最初にこのフェーズを完了させ、結果を報告してからユーザー承認を待つ。

1. `C:\DEV\SnsManagementTool` の存在確認とディレクトリ構成の把握
2. `requirements.txt` / `package.json` / 起動スクリプトの特定
3. 使用ポート番号の特定（コード内 `PORT`, `app.run(port=...)` 等をgrep）
4. **予定管理システム(ポート5000/タスク名 ScheduleServer)との競合可能性を確認**
5. 運用端末の空きディスク容量と Python バージョン確認

**Claude Code への指示例**:
> フェーズAの調査を実施し、結果をMarkdown表で報告してください。ユーザーが承認するまで次のフェーズに進まないでください。

---

### フェーズ B — バックアップ作成（Git代替のバージョン管理）

**このプロジェクトはGit未使用のため、フォルダバックアップが唯一のバージョン履歴となる。**

1. バックアップフォルダ命名規則
   - 形式: `SnsManagementTool_v{バージョン}_{YYYYMMDD_HHMM}`
   - 例: `SnsManagementTool_v1.0.0_20260911_1030`

2. 実行コマンド（Claude Code に生成させる）

```cmd
robocopy "C:\DEV\SnsManagementTool" "D:\Backup\SnsManagementTool\SnsManagementTool_v1.0.0_20260911_1030" /E /XD __pycache__ node_modules .venv
```

3. バックアップ後に **`VERSION.txt`** をバックアップフォルダ直下に作成

```
バージョン: 1.0.0
バックアップ日時: 2026-09-11 10:30
作成者: 新名 洋幸
変更概要: 初回リリース／機能追加内容の一行要約
リリース種別: 新規リリース ／ 更新 ／ ホットフィックス
```

4. 過去バックアップの保持ルール
   - 直近5世代は必ず保持
   - それより古いものは月次アーカイブ(`_archive/`)へ移動

---

### フェーズ C — 本番反映（初回リリース）

初回のみ実行。既に本番配置済みなら **フェーズ D** を使用する。

1. Python 環境確認 (`python --version` が 3.9 以上)
2. 配置先フォルダ `C:\Apps\SnsManagementTool` 作成
3. robocopy でソース配置（除外: `.git`, `__pycache__`, `node_modules`, `.venv`, 開発用DB, 開発用鍵）

```cmd
robocopy "C:\DEV\SnsManagementTool" "C:\Apps\SnsManagementTool" /E /XD .git __pycache__ .venv node_modules /XF *.db .secret_key .env
```

4. 依存パッケージインストール

```cmd
cd C:\Apps\SnsManagementTool
pip install -r requirements.txt
```

5. 起動確認（フォアグラウンドで手動起動 → ブラウザで動作確認 → Ctrl+C 停止）
6. タスクスケジューラ登録（**予定管理システムと別タスク名で登録**）

```cmd
schtasks /Create /TN "SnsManagementServer" /TR "C:\Apps\SnsManagementTool\start_server.bat" /SC ONSTART /RU SYSTEM /RL HIGHEST /F
```

7. ファイアウォール開放（例: ポート5010）

```cmd
netsh advfirewall firewall add rule name="SNS Management Server" dir=in action=allow protocol=tcp localport=5010
```

---

### フェーズ D — 本番更新（2回目以降）

1. **必ずフェーズBのバックアップを取得済みであること**
2. サーバー停止

```cmd
schtasks /End /TN "SnsManagementServer"
```

3. ソースコード同期（DB・鍵・ログは保護）

```cmd
robocopy "C:\DEV\SnsManagementTool" "C:\Apps\SnsManagementTool" /E /PURGE /XD __pycache__ .git .venv node_modules /XF *.db .secret_key .env *.log
```

4. 依存追加の有無確認

```cmd
cd C:\Apps\SnsManagementTool
pip install -r requirements.txt --quiet
```

5. サーバー再起動

```cmd
schtasks /Run /TN "SnsManagementServer"
```

6. 動作確認（ブラウザで `http://localhost:5010`）

---

### フェーズ E — リリース後確認とログ記録

Claude Code に必ず実施させる。

1. サーバー稼働確認: `schtasks /Query /TN "SnsManagementServer"`
2. **予定管理システムが停止していないこと** も併せて確認: `schtasks /Query /TN "ScheduleServer"`
3. ログファイル(`db/service_stdout.log` 等)にエラーが出ていないか確認
4. `docs/リリース履歴.md` に以下を追記

```
## v1.0.0 - 2026-09-11
- 実施者: 新名 洋幸
- 種別: 新規リリース
- 変更概要: 初回リリース
- バックアップ: D:\Backup\SnsManagementTool\SnsManagementTool_v1.0.0_20260911_1030
- 確認結果: 正常稼働
```

---

## 3. Claude Code に禁止する操作

CLAUDE.md 準拠。SNS管理ツールでも同ルール。

- `.env` / `*.pem` / `*.key` / `secrets.json` の読み書き
- `curl` / `wget` / `scp` / `ssh` によるデータ外部送信
- `rm -rf` / `del /S /Q` を **本番フォルダに** 実行
- 予定管理システム(ポート5000, タスク名 ScheduleServer)への干渉
- バックアップ未取得のままの本番上書き
- `pip install` / `npm install` の無確認実行

---

## 4. 想定Q&A（Claude Code が判断に迷った場合の指針）

| 状況 | Claude Code の取るべき行動 |
|------|---------------------------|
| ポート競合を検知 | 作業停止しユーザーに別ポート番号を確認 |
| バックアップ先の容量不足 | 作業停止し古いバックアップの整理をユーザーへ依頼 |
| requirements.txt に破壊的変更を検知 | フェーズC/Dを止めユーザーに影響範囲を報告 |
| 本番DBファイルの上書きリスクを検知 | 即時停止しユーザー承認を得るまで再開しない |
| 予定管理システムのプロセスが落ちた | 直ちに停止しユーザーへ通知、SNS作業より復旧優先 |

---

## 5. 参考リンク

- 一般手順書: `02_一般的なリリース手順書.md`
- 保守/バージョン管理: `03_リリース後メンテナンスとバージョン管理規定.md`
- 予定管理システム: `C:\DEV\ScheduleManagement\deploy\RELEASE_GUIDE.md`

---

以上を Claude Code への **SNS管理ツール リリース指示書** とする。
