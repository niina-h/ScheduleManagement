# DB バックアップ・復旧手順書

予定管理システムの SQLite DB (`web_app.db`) を、端末故障時に復旧できるようにするための手順書。

> 最終更新：2026-05-18 / バージョン：v1.11.0

---

## 1. バックアップ構成（3層）

| 層 | 場所 | 用途 | 実行主体 | 頻度 |
|---|---|---|---|---|
| ① 本番ローカル世代 | `\\192.168.70.141\C$\App\ScheduleManagement\db\backups\` | 直近の誤操作復旧 | 本番サーバーのタスクスケジューラ | 毎日 18:00 |
| ② オフサイト | `c:\DEV(ClaudCode)\Backup\ScheduleManagement\` | 端末故障時の復旧 | 開発機のタスクスケジューラ | 毎日 18:30 |
| ③ 手動取得 | マスタ権限ユーザーが任意でダウンロード | 確認・予備 | マスタ操作 | 任意 |

- ファイル名：`web_app_YYYYMMDD_HHMMSS.db`
- 各層 **最新30世代を自動保持**（古いものは自動削除）

---

## 2. 初期設定（最初に1回だけ実施）

### 2-1. 本番サーバー側にタスクスケジューラを登録

1. 本番サーバー（192.168.70.141）にリモートデスクトップで接続
2. `C:\App\ScheduleManagement\bat\register_backup_task.bat` を **管理者として実行**
3. 「タスク登録完了」と表示されれば OK
4. 確認：`schtasks /Query /TN "ScheduleBackup"`

### 2-2. 開発機側（このマシン）にタスクスケジューラを登録

1. このマシンで `c:\DEV(ClaudCode)\ScheduleManagement\bat\register_pull_backup_task.bat` を **管理者として実行**
2. 「タスク登録完了」と表示されれば OK
3. 保存先：`c:\DEV(ClaudCode)\Backup\ScheduleManagement\`
4. 確認：`schtasks /Query /TN "ScheduleBackupPull"`

### 2-3. 動作確認（手動実行）

各バッチを一度手動実行し、バックアップファイルが作成されることを確認：

```cmd
REM 本番サーバーで実行
C:\App\ScheduleManagement\bat\backup_db.bat

REM 開発機で実行
c:\DEV(ClaudCode)\ScheduleManagement\bat\pull_backup_from_prod.bat
```

成功時：`db\backups\` または `c:\DEV(ClaudCode)\Backup\ScheduleManagement\` に `web_app_YYYYMMDD_HHMMSS.db` が出来ます。

---

## 3. 日常運用

### 3-1. 自動実行（推奨）
タスクスケジューラ登録後は **何もしなくて OK**。毎日18:00（本番）と18:30（開発機）に自動でバックアップ。

### 3-2. ログ確認
万一バックアップが失敗していないかは、以下のログで確認：
- 本番側：`\\192.168.70.141\C$\App\ScheduleManagement\logs\backup.log`
- 開発機側：`c:\DEV(ClaudCode)\Backup\ScheduleManagement\backup.log`

### 3-3. 手動ダウンロード（マスタ権限）
ブラウザの管理者ダッシュボード右上の **「💾 DBバックアップ取得」** ボタンで、現時点の DB ファイルをダウンロード可能。

---

## 4. 復旧手順

### ケースA：本番サーバー（192.168.70.141）が故障した場合

**前提**：開発機に最新のオフサイトバックアップが残っている。

1. **新しい本番環境を準備**（同じ Windows 環境）
2. プログラム一式をデプロイ：
   - `deploy/install.bat` などのインストーラを実行
   - または `bat/deploy_code.bat` で開発機からコピー
3. **オフサイトバックアップから DB を復元**：
   - `c:\DEV(ClaudCode)\Backup\ScheduleManagement\` から最新の `web_app_YYYYMMDD_HHMMSS.db` を選ぶ
   - 新本番サーバーの `C:\App\ScheduleManagement\db\web_app.db` にコピー（リネーム）
4. サーバーを起動：`bat/start_server.bat` または タスクスケジューラ `ScheduleServer` を有効化
5. ブラウザでアクセスして動作確認

### ケースB：誤操作で DB の中身が壊れた場合

**前提**：本番サーバー自体は生きていて、`db\backups\` に世代がある。

1. 本番サーバーを停止：`bat/server_stop.bat`
2. `\\192.168.70.141\C$\App\ScheduleManagement\db\backups\` から **誤操作前の時点** の `web_app_YYYYMMDD_HHMMSS.db` を選ぶ
3. 現在の `web_app.db` を念のため `web_app_BROKEN.db` 等にリネームして退避
4. 選んだバックアップを `web_app.db` にリネームしてコピー
5. サーバーを起動：`bat/server_start.bat`

### ケースC：開発機も本番も同時に壊れた場合（最悪ケース）

**前提**：マスタ権限ユーザーが定期的に「💾 DBバックアップ取得」で手動ダウンロードしていた。

1. マスタユーザーのPCから手動ダウンロード済みの `web_app_YYYYMMDD_HHMMSS.db` を取り出す
2. ケースA と同様に新環境を立ち上げ、DBファイルを `web_app.db` として配置

---

## 5. 注意事項

- **SQLite ファイルは Flask + waitress が開いた状態でもコピー可能**ですが、コピー中の整合性を確実にするには、極短い時間で書き込みのないタイミングでバックアップが望ましい。タスクスケジューラの18:00 は業務終了後を想定。
- バックアップファイルには **個人情報・組織情報** が含まれます。社外への持ち出しは慎重に。
- 「💾 DBバックアップ取得」はマスタ権限のみ（403 制御済み）。

---

## 6. 削除順序（世代管理）

各層とも、ファイル更新日時降順で **31番目以降** を自動削除します。古いバックアップを残したい場合は、別のフォルダに移動してから次回バックアップを待つこと。
