# 予定管理システム — プログラム更新手順書

## 概要

開発PCでプログラムを修正した後、本番環境に反映する手順。
**データベース（DB）の内容は一切変更しない。**

---

## 更新方法（自動）

### update.bat による一括更新

`deploy\update.bat` をダブルクリックするだけで完了。

```
[1/4] サーバー停止（タスクスケジューラ経由）
[2/4] ファイル同期（ソースコードのみ）
[3/4] パッケージ確認（新しい依存があれば追加）
[4/4] サーバー再起動
```

#### 事前設定

初回のみ、`deploy\update.bat` のパスを環境に合わせて編集する：

```batch
set SOURCE=c:\DEV(ClaudCode)\ScheduleManagement   ← 開発元のパス
set DEST=C:\Apps\ScheduleManagement                ← 本番先のパス
```

---

## 更新対象と保護対象

### 上書き更新されるファイル

| 対象 | 内容 |
|------|------|
| `web_app/` | アプリケーション全体（Python・テンプレート・CSS） |
| `data/` | 初期ユーザーデータ |
| `deploy/` | デプロイ用バッチファイル |
| `run_web.py` | 開発用起動スクリプト |
| `run_production.py` | 本番用起動スクリプト |
| `requirements_web.txt` | 依存パッケージ定義 |

### 上書きされないファイル（保護対象）

| 対象 | 理由 |
|------|------|
| `db/web_app.db` | 本番データベース（ユーザーデータ・予定データ全て） |
| `db/.secret_key` | セッション暗号化キー（変わるとログイン中ユーザーが全員ログアウトされる） |
| `db/service_stdout.log` | サーバーログ |
| `db/service_stderr.log` | エラーログ |

---

## 更新方法（手動）

update.bat が使えない場合の手動手順。

### 手順1: サーバーを停止

```cmd
schtasks /End /TN "ScheduleServer"
```

3秒ほど待つ。

### 手順2: ファイルをコピー

```cmd
robocopy "開発元\web_app" "C:\Apps\ScheduleManagement\web_app" /E /PURGE /XD __pycache__
robocopy "開発元\data" "C:\Apps\ScheduleManagement\data" /E
robocopy "開発元\deploy" "C:\Apps\ScheduleManagement\deploy" /E
copy /Y "開発元\run_production.py" "C:\Apps\ScheduleManagement\run_production.py"
copy /Y "開発元\requirements_web.txt" "C:\Apps\ScheduleManagement\requirements_web.txt"
```

### 手順3: 依存パッケージ更新（必要な場合）

```cmd
cd C:\Apps\ScheduleManagement
pip install -r requirements_web.txt --quiet
```

### 手順4: サーバーを再起動

```cmd
schtasks /Run /TN "ScheduleServer"
```

---

## DBスキーマ変更を伴う更新の場合

通常の更新はソースコードのみだが、DBテーブル構造が変わる場合は追加手順が必要。

### 手順

1. **事前にDBをバックアップ**

```cmd
copy "C:\Apps\ScheduleManagement\db\web_app.db" "C:\Apps\ScheduleManagement\db\web_app_%date:~0,4%%date:~5,2%%date:~8,2%.db.bak"
```

2. 通常の更新手順を実行（update.bat）
3. アプリ起動時に `database.py` の `init_db()` が自動でスキーマを更新
4. 動作確認後、バックアップファイルは1週間程度保持してから削除

---

## Excel テンプレートの更新

`reports/tpl/` 配下のExcelテンプレートを更新する場合：

```cmd
robocopy "開発元\reports\tpl" "C:\Apps\ScheduleManagement\reports\tpl" /E
```

> テンプレート更新はサーバー再起動不要（リクエスト毎にファイルを読み込むため）。

---

## 更新後の確認事項

| # | 確認内容 | 方法 |
|---|---------|------|
| 1 | サーバーが起動しているか | `schtasks /Query /TN "ScheduleServer"` で「実行中」を確認 |
| 2 | ログインページが表示されるか | ブラウザで http://（運用PC）:5000 にアクセス |
| 3 | 既存データが保持されているか | ログインして週間予定・日次実績が従来通り表示されること |
| 4 | 新機能が動作するか | 更新内容に応じて該当機能を確認 |

---

## ロールバック（更新を戻す場合）

問題が発生した場合、以前のバージョンに戻す。

### 手順

1. サーバーを停止

```cmd
schtasks /End /TN "ScheduleServer"
```

2. 問題のあるファイルを以前のバージョンで上書き
   - Git 管理している場合: 開発PCで `git checkout 前のコミット` → 再度 update.bat
   - Git 未管理の場合: バックアップから復元

3. サーバーを再起動

```cmd
schtasks /Run /TN "ScheduleServer"
```

> **注意**: DB のロールバックが必要な場合は、手順「DBスキーマ変更を伴う更新」で作成したバックアップファイルから復元する。
