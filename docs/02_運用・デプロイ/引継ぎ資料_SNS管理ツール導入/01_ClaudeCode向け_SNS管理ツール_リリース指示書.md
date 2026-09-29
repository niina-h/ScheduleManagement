# Claude Code 向け指示書 — SNS管理ツール 更新作業

本ドキュメントは、本番PC（D2KMXR3）で **既に稼働中** の SNS管理ツール（SNS Manager Studio）を
更新する際、Claude Code に対して与える標準指示書である。

**最終更新：2026-09-28（本番実態にあわせて全面改訂）**

---

## ★ 改訂履歴

| 日付 | 内容 |
|------|------|
| 2026-09-14 | 初版作成（新規リリース前提） |
| 2026-09-28 | **全面改訂**。実機調査により SNS管理ツールは既に本番稼働中と判明。新規リリース手順から更新作業手順へ変更。予定管理システムの本番配置も `C:\DEV\` → `C:\App\` に訂正 |

---

## 0. 本番環境の実測値（2026-09-28 実機確認済み）

### 0.1 運用PC

| 項目 | 実測値 | 確認方法 |
|------|--------|---------|
| コンピュータ名 | **D2KMXR3** | `hostname` |
| IPアドレス | **192.168.70.141 / 24** | `ipconfig` |
| OS | **Windows 11 Pro 10.0.26200** | `winver` |
| ドメイン | **crown-commune.local** | `whoami /fqdn` |
| ログオンアカウント | **niina-h** | `whoami` |
| Python（本番共用） | **C:\Python\Python313\python.exe**（3.13.3） | `where python` |
| pip プロキシ | **不要（直接）** | 実績 |
| ストレージ | 512GB SSD、Cドライブのみ（**Dドライブなし**） | `Get-PSDrive` |
| C空き容量 | 約 203GB | `Get-PSDrive C` |
| SMB | TCP 445 / 139 LISTENING | `netstat` |

### 0.2 予定管理システム（★干渉禁止・稼働中）

| 項目 | 実測値 |
|------|--------|
| 本番配置 | **C:\App\ScheduleManagement\** |
| 本番DB | **C:\App\ScheduleManagement\db\web_app.db** |
| ポート | **TCP 5000** |
| タスク名 | **ScheduleServer** |
| 実行 | `pythonw.exe run_production.py`（waitress） |
| 実行ユーザー | SYSTEM（ログイン不要） |
| トリガー | システム起動時 |
| 稼働状態 | **稼働中・業務時間中は利用者が接続** |

> `C:\DEV\ScheduleManagement\` は **開発用フォルダ**であり本番ではない。混同しないこと。

### 0.3 SNS管理ツール（★更新対象・稼働中）

| 項目 | 実測値 |
|------|--------|
| 本番配置 | **C:\App\SnsManagementTool\** |
| 実装 | **Streamlit**（`app.py`） |
| 起動コマンド | `python -m streamlit run app.py --server.port 5010 --server.address 0.0.0.0 --server.headless true` |
| 起動スクリプト | `C:\App\SnsManagementTool\start_server.bat` |
| 再起動スクリプト | `C:\App\SnsManagementTool\restart_server.bat` |
| ポート | **TCP 5010** |
| タスク名 | **SnsManagementServer** |
| 実行ユーザー | SYSTEM（ログイン不要） |
| トリガー | システム起動時（1分遅延） |
| 稼働状態 | **稼働中** |
| 設定ファイル | `keyword_groups.json`、`.env`（APIキー） |
| 依存 | `requirements.txt`（Streamlit 等） |
| バージョン管理 | **Git 未使用**（フォルダバックアップ） |

### 0.4 バックアップ先

| 種別 | パス | 役割 |
|------|------|------|
| **A. NAS 正本** | `\\crown-commune\cloud\Data\商品開発室\50 DEV\54SNS管理ツール\05_開発\03_ソースコード・成果物\v{バージョン}\` | 履歴の正本・永年保管 |
| **B. ローカル** | `C:\Backup\SnsManagementTool\` | ロールバック用・直近5世代 |

> **注意**：2026-09-28 時点で本端末から NAS への `Test-Path` が False。
> 更新作業前に疎通を確認すること（DNS／権限／パス綴りの調査が必要）。

### 0.5 連絡先

| 役割 | 氏名 | 連絡先 |
|------|------|--------|
| リリース責任者 | 新名 洋幸 | niina-h@crown-c.com |
| 運用PC管理者 | 新名 洋幸 | niina-h@crown-c.com |
| 予定管理システム保守 | 新名 洋幸 | niina-h@crown-c.com |
| SNS管理ツール開発担当 | `___________________` | `___________________` |
| 障害時二次エスカレーション | `___________________` | `___________________` |

### 0.6 更新作業時に確認が必要な項目

| # | 項目 | 確認先 |
|---|------|--------|
| 1 | 開発元パス（開発PC上のフォルダ） | 開発担当 |
| 2 | 更新するバージョン番号・変更内容 | 開発担当 |
| 3 | `requirements.txt` の変更有無 | 開発担当 |
| 4 | NAS 疎通問題の解決 | NAS管理者 |
| 5 | 実施日時 | 運用担当 |

---

## 1. Claude Code への指示テンプレート

以下をコピーし、空欄を埋めてプロンプトとして使用する。

```
# 依頼内容
本番PC で稼働中の SNS管理ツール（SNS Manager Studio）を更新してください。
以下の手順に沿って、各フェーズ終了時に確認を取りながら段階的に進めてください。

## 参考資料
- 本指示書（01_ClaudeCode向け_SNS管理ツール_リリース指示書.md）
- 運用手順（05_SNS管理ツール運用手順.md）

## 実行環境情報
- 運用PC: D2KMXR3 / 192.168.70.141
- SNS管理ツール本番配置: C:\App\SnsManagementTool\
- ポート: 5010 / タスク名: SnsManagementServer
- 実装: Streamlit（app.py）
- Python: C:\Python\Python313\python.exe（3.13.3、予定管理システムと共用）
- バックアップA（NAS正本）: \\crown-commune\cloud\Data\商品開発室\50 DEV\54SNS管理ツール\05_開発\03_ソースコード・成果物\v{バージョン}\
- バックアップB（ローカル）: C:\Backup\SnsManagementTool\
- 更新元パス: ___________________
- 新バージョン番号: ___________________

## 絶対に守ること
- 予定管理システム（ポート5000 / タスク ScheduleServer / C:\App\ScheduleManagement\）を止めない
- .env / APIキー / keyword_groups.json は読み書きしない・上書きしない
- outputs\ は上書きしない
- バックアップ A・B の両方を取得してから更新する（A が取れなければ中断）
- バックアップに .venv を含めない
- 破壊的操作は必ず事前確認
- 各フェーズ終了後に「次に進んでよいか」確認する
```

---

## 2. 作業フロー（フェーズ別）

### フェーズ A — 事前調査（読み取りのみ）

結果を報告し、承認を待ってから次へ進む。

1. 両システムの稼働確認

```cmd
netstat -ano | findstr :5000
netstat -ano | findstr :5010
```

2. 更新元フォルダの構成確認（`app.py`、`requirements.txt` の有無）
3. 現行本番との差分確認（変更ファイルの洗い出し）
4. `requirements.txt` の差分有無を確認
5. NAS 疎通確認

```cmd
powershell -Command "Test-Path '\\crown-commune\cloud\Data\商品開発室\50 DEV\54SNS管理ツール\05_開発\03_ソースコード・成果物'"
```

6. Cドライブ空き容量確認

**NAS 疎通が False の場合は、この時点で作業を中断しユーザーに報告する。**

---

### フェーズ B — バックアップ（A・B 両方必須）

#### A. NAS リリース版アーカイブ（正本）

```cmd
robocopy "C:\App\SnsManagementTool" ^
         "\\crown-commune\cloud\Data\商品開発室\50 DEV\54SNS管理ツール\05_開発\03_ソースコード・成果物\v{新バージョン}" ^
         /E /XD .venv __pycache__ .git node_modules outputs /XF *.log
```

除外：`.venv`（★必須）、`__pycache__`、`.git`、`outputs`、`*.log`

#### B. ローカルバックアップ

```cmd
robocopy "C:\App\SnsManagementTool" ^
         "C:\Backup\SnsManagementTool\SnsManagementTool_v{新バージョン}_{YYYYMMDD_HHMM}" ^
         /E /XD .venv __pycache__ outputs
```

#### VERSION.txt（A・B 両方に作成）

```
============================================
プロジェクト: SNS管理ツール（SNS Manager Studio）
バージョン: {新バージョン}
リリース日時: {YYYY-MM-DD HH:MM}
作成者: {氏名}
リリース種別: {メジャー／マイナー／パッチ／ホットフィックス}
============================================

【変更概要】
- {箇条書き}

【依存パッケージ変更】
- あり／なし（ある場合は差分）

【除外物】
.venv, __pycache__, .git, outputs, *.log

【本番配置先】
C:\App\SnsManagementTool

【運用PC】
D2KMXR3 (192.168.70.141)
```

#### 検証

- A のフォルダサイズが妥当か（`.venv` 混入で肥大化していないか）
- `VERSION.txt` が A・B 双方に存在するか

**片方でも失敗したらフェーズ C に進まない。**

---

### フェーズ C — 更新反映

1. サーバー停止

```cmd
schtasks /End /TN "SnsManagementServer"
```

2. ファイル同期（設定・出力物は保護）

```cmd
robocopy "{更新元パス}" "C:\App\SnsManagementTool" ^
         /E /XD .venv __pycache__ .git outputs ^
         /XF .env keyword_groups.json *.log
```

**保護対象（上書き禁止）**：`.env`、`keyword_groups.json`、`outputs\`、`*.log`

3. 依存パッケージ更新（`requirements.txt` に変更がある場合のみ）

```cmd
cd C:\App\SnsManagementTool
C:\Python\Python313\python.exe -m pip install -r requirements.txt
```

> **注意**：Python は予定管理システムと共用。パッケージ競合が起きると予定管理システムにも影響する。
> バージョン指定の変更がある場合は、事前にユーザーへ影響範囲を報告して承認を得る。

4. サーバー起動

```cmd
schtasks /Run /TN "SnsManagementServer"
```

Streamlit の起動には数十秒かかる場合がある。

---

### フェーズ D — 確認

1. 両システムのポート確認

```cmd
netstat -ano | findstr :5010
netstat -ano | findstr :5000
```

**両方 LISTENING であること。**

2. HTTP 応答確認

```cmd
powershell -Command "(Invoke-WebRequest -Uri 'http://localhost:5010/' -UseBasicParsing -TimeoutSec 15).StatusCode"
powershell -Command "(Invoke-WebRequest -Uri 'http://localhost:5000/' -UseBasicParsing -TimeoutSec 15).StatusCode"
```

3. 重複プロセスの確認（旧プロセスが残っていないか）

```cmd
powershell -Command "Get-Process python,pythonw -ErrorAction SilentlyContinue | Select-Object Id,StartTime"
```

4. 機能確認（更新内容に応じた画面操作）

5. リリース履歴に追記

```
## v{バージョン} - {YYYY-MM-DD}
- 実施者: {氏名}
- 種別: {リリース種別}
- 変更概要: {内容}
- バックアップA: \\crown-commune\...\v{バージョン}
- バックアップB: C:\Backup\SnsManagementTool\SnsManagementTool_v{バージョン}_{日時}
- 確認結果: 正常稼働
```

---

### フェーズ E — ロールバック（問題発生時）

```cmd
schtasks /End /TN "SnsManagementServer"

robocopy "C:\Backup\SnsManagementTool\{バックアップフォルダ}" ^
         "C:\App\SnsManagementTool" ^
         /E /XD .venv __pycache__ outputs /XF .env keyword_groups.json

schtasks /Run /TN "SnsManagementServer"
```

ローカルバックアップが使えない場合は NAS の該当バージョンから復元する。

---

## 3. Claude Code に禁止する操作

- `.env` / `*.pem` / `*.key` / APIキーを含むファイルの読み書き
- `keyword_groups.json`、`outputs\` の上書き
- `curl` / `wget` / `scp` / `ssh` によるデータ外部送信
- 本番フォルダへの `rm -rf` / `del /S /Q`
- **予定管理システム（ポート5000 / `ScheduleServer` / `C:\App\ScheduleManagement\`）への干渉**
- `Get-Process python | Stop-Process` のような両システムを巻き込む操作
- バックアップ A（NAS）未取得のままの更新
- バックアップへの `.venv` 混入
- `pip install` の無確認実行（共用 Python のため影響大）
- Python 本体のバージョン変更

---

## 4. 判断に迷った場合の指針

| 状況 | Claude Code の取るべき行動 |
|------|---------------------------|
| NAS にアクセス不可 | 作業中断。ローカルB のみでの続行は不可 |
| バックアップに `.venv` が混入 | A フォルダを削除し、除外を確認して再取得 |
| `requirements.txt` に破壊的変更 | 更新を止め、予定管理システムへの影響をユーザーに報告 |
| ポート5010 が LISTENING にならない | 1分待って再確認。それでも駄目なら `start_server.bat` を直接実行してエラーを確認 |
| 予定管理システムが停止した | **直ちに作業中断。`schtasks /Run /TN "ScheduleServer"` で復旧を最優先**。SNS作業は後回し |
| python プロセスが想定外に複数ある | 停止せず、PID と起動時刻をユーザーに報告して指示を仰ぐ |
| 業務時間中に更新を依頼された | 利用者影響を報告し、時間帯の再検討を提案 |

---

## 5. 参考リンク

- 運用手順：`05_SNS管理ツール運用手順.md`
- 本番PC構成図：`04_本番PCシステム概要図.html`
- 一般的なリリース手順：`02_一般的なリリース手順書.md` / `.html`
- バージョン管理規定：`03_リリース後メンテナンスとバージョン管理規定.md` / `.html`
- 本体セットアップ手順：`C:\App\SnsManagementTool\setup_instructions.txt`
- 予定管理システム側の手順：`../既存資料_予定管理システム/デプロイ手順書.md`

---

## 6. 作業開始前チェックリスト

- [ ] §0.6 の確認事項（更新元パス・バージョン番号・依存変更）が確定している
- [ ] NAS への疎通が確認できている
- [ ] 実施日時が業務影響の少ない時間帯である
- [ ] §1 指示テンプレートの空欄を埋めた
- [ ] 予定管理システムが稼働中であることを確認した（止めないため）

**未確定項目がある状態での作業着手は禁止。**
