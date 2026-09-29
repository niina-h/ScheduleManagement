# SNS管理ツール 運用手順書

本番PC（D2KMXR3 / 192.168.70.141）で稼働中の **SNS Manager Studio（SNS管理ツール）** の運用手順。

最終更新：2026-09-28

---

## 1. システム構成

| 項目 | 値 |
|------|-----|
| アプリ名 | SNS Manager Studio |
| 実装 | Streamlit（Python） |
| 本番配置 | `C:\App\SnsManagementTool\` |
| エントリポイント | `app.py` |
| ポート | TCP 5010 |
| タスク名 | `SnsManagementServer` |
| 実行ユーザー | SYSTEM（ユーザーログイン不要） |
| トリガー | システム起動時（1分遅延） |
| Python | `C:\Python\Python313\python.exe`（予定管理システムと共用） |
| アクセスURL | http://192.168.70.141:5010 |

### 主要フォルダ

```
C:\App\SnsManagementTool\
├─ app.py                  # Streamlit エントリポイント
├─ main.py
├─ config.py
├─ requirements.txt        # 依存パッケージ
├─ keyword_groups.json     # キーワード設定
├─ start_server.bat        # 起動スクリプト（タスクから呼ばれる）
├─ restart_server.bat      # 再起動スクリプト
├─ setup_instructions.txt  # 初期セットアップ手順
├─ SnsManagementServer.xml # タスク定義のエクスポート
├─ modules\                # 機能モジュール群
├─ ui\                     # 画面定義
├─ scripts\                # 検証用スクリプト
└─ outputs\                # 出力ファイル
```

### 同居システムとの分離

| | 予定管理システム | SNS管理ツール |
|---|---|---|
| ポート | 5000 | **5010** |
| タスク名 | `ScheduleServer` | **`SnsManagementServer`** |
| フォルダ | `C:\App\ScheduleManagement\` | **`C:\App\SnsManagementTool\`** |

ポート・タスク名・フォルダが全て別なので、片方の操作がもう片方に影響することはない。
ただし **Python 本体・タスクスケジューラ・ディスク・ネットワークは共用**。

---

## 2. 日常操作

### 2.1 稼働確認

```cmd
netstat -ano | findstr :5010
```

LISTENING が表示されれば稼働中。ブラウザで http://192.168.70.141:5010 を開いて画面が出ることも確認する。

### 2.2 再起動

```cmd
C:\App\SnsManagementTool\restart_server.bat
```

または手動で：

```cmd
schtasks /End /TN "SnsManagementServer"
schtasks /Run /TN "SnsManagementServer"
```

Streamlit の起動には数十秒かかる場合がある。再起動後は `netstat` で LISTENING を確認する。

### 2.3 停止・起動

```cmd
schtasks /End /TN "SnsManagementServer"    # 停止
schtasks /Run /TN "SnsManagementServer"    # 起動
```

### 2.4 状態照会（管理者権限が必要）

```cmd
schtasks /Query /TN "SnsManagementServer" /V /FO LIST
```

---

## 3. 更新手順（アプリのコードを変えるとき）

### 3.1 事前準備

1. **予定管理システムが稼働中であることを確認**（止めないため）

```cmd
netstat -ano | findstr :5000
```

2. 業務時間外に実施するのが望ましい（SNS管理ツール利用者への影響を避ける）

### 3.2 バックアップ（2階層・両方必須）

#### A. NAS リリース版アーカイブ（正本・必須）

```cmd
robocopy "C:\App\SnsManagementTool" ^
         "\\crown-commune\cloud\Data\商品開発室\50 DEV\54SNS管理ツール\05_開発\03_ソースコード・成果物\v1.1.0" ^
         /E /XD .venv __pycache__ .git node_modules outputs /XF *.log
```

- **`.venv` は必ず除外**（環境依存・容量肥大）
- `outputs\` も除外（生成物のため）
- バージョンフォルダ直下に `VERSION.txt` を作成する（§3.5 参照）

#### B. ローカルバックアップ（ロールバック用）

```cmd
robocopy "C:\App\SnsManagementTool" ^
         "C:\Backup\SnsManagementTool\SnsManagementTool_v1.1.0_20260928_1800" ^
         /E /XD .venv __pycache__ outputs
```

> **A（NAS）が未取得のまま更新してはならない。** 正本が残らず、後日の復元・監査ができなくなる。

### 3.3 サーバー停止

```cmd
schtasks /End /TN "SnsManagementServer"
```

### 3.4 ファイル反映

開発元から本番へコピーする。**設定ファイル・出力物は上書きしない。**

```cmd
robocopy "<開発元パス>" "C:\App\SnsManagementTool" ^
         /E /XD .venv __pycache__ .git outputs ^
         /XF .env keyword_groups.json *.log
```

#### 保護対象（上書き禁止）

| ファイル／フォルダ | 理由 |
|------------------|------|
| `.env` | API キー等の本番設定 |
| `keyword_groups.json` | 運用中のキーワード設定 |
| `outputs\` | 生成済みの出力物 |
| `*.log` | 障害調査に必要 |

### 3.5 依存パッケージの更新（`requirements.txt` が変わった場合のみ）

```cmd
cd C:\App\SnsManagementTool
C:\Python\Python313\python.exe -m pip install -r requirements.txt
```

> **注意**：Python は予定管理システムと共用。パッケージのバージョン競合が起きると
> 予定管理システム側にも影響する。更新前に差分を確認すること。

### 3.6 サーバー起動・確認

```cmd
schtasks /Run /TN "SnsManagementServer"
```

数十秒待ってから：

```cmd
netstat -ano | findstr :5010
netstat -ano | findstr :5000
```

**両方 LISTENING であることを確認する。** ブラウザで両システムにアクセスして動作確認する。

### 3.7 記録

`VERSION.txt` をバックアップ先（A・B 両方）に作成する。

```
============================================
プロジェクト: SNS管理ツール（SNS Manager Studio）
バージョン: 1.1.0
リリース日時: 2026-09-28 18:00
作成者: (氏名)
リリース種別: マイナーリリース
============================================

【変更概要】
- (箇条書き)

【依存パッケージ変更】
- あり／なし（ある場合は差分）

【除外物】
.venv, __pycache__, .git, outputs, *.log, .env, keyword_groups.json

【本番配置先】
C:\App\SnsManagementTool

【運用PC】
D2KMXR3 (192.168.70.141)
```

---

## 4. ロールバック

更新後に問題が発生した場合。

```cmd
schtasks /End /TN "SnsManagementServer"

robocopy "C:\Backup\SnsManagementTool\SnsManagementTool_v1.0.0_20260911_1030" ^
         "C:\App\SnsManagementTool" ^
         /E /XD .venv __pycache__ outputs /XF .env keyword_groups.json

schtasks /Run /TN "SnsManagementServer"
```

ローカルバックアップが無い／壊れている場合は NAS の該当バージョンフォルダから復元する。

---

## 5. トラブルシューティング

| 症状 | 確認・対処 |
|------|-----------|
| ポート5010 が LISTENING にならない | タスクが動いているか `schtasks /Query`。Streamlit の起動に時間がかかるため1分程度待つ |
| ブラウザから接続できない | ファイアウォールを確認：`netsh advfirewall firewall show rule name="SNS Management Server"` |
| 起動直後に落ちる | `start_server.bat` をコマンドプロンプト内で直接実行してエラーを確認（ダブルクリック不可） |
| 依存パッケージのエラー | `C:\Python\Python313\python.exe -m pip install -r requirements.txt` を再実行 |
| pythonw/python が2つ以上ある | 旧プロセスが残っている。PID を確認して古い方を停止 |
| 予定管理システムが停止した | **最優先で復旧**：`schtasks /Run /TN "ScheduleServer"` |

### ファイアウォール開放（未設定の場合）

```cmd
netsh advfirewall firewall add rule name="SNS Management Server" dir=in action=allow protocol=tcp localport=5010
```

---

## 6. 禁止事項

- 予定管理システム（ポート5000 / `ScheduleServer`）の停止・干渉
- `Get-Process python | Stop-Process` のような **両システムを巻き込む操作**
- NAS リリース版アーカイブ未取得のままの更新
- バックアップに `.venv` を含める
- `.env` / API キーを含むファイルの外部送信・別端末コピー
- Python 本体のバージョン変更（両システムに影響。実施する場合は両方の依存を事前検証）

---

## 7. 定期メンテナンス

| 頻度 | 項目 |
|------|------|
| 毎日 | 稼働確認（5000・5010 の両方） |
| 週次 | ディスク空き容量（20%以上維持）、`outputs\` の肥大化確認 |
| 月次 | 依存パッケージの脆弱性確認、ローカルバックアップの世代整理（直近5世代） |
| 四半期 | API キーの有効期限・利用量確認 |

---

## 8. 参考

- 本体のセットアップ手順：`C:\App\SnsManagementTool\setup_instructions.txt`
- タスク定義：`C:\App\SnsManagementTool\SnsManagementServer.xml`
- システム構成図：[04_本番PCシステム概要図.html](04_本番PCシステム概要図.html)
- バージョン管理規定：[03_リリース後メンテナンスとバージョン管理規定.md](03_リリース後メンテナンスとバージョン管理規定.md)
