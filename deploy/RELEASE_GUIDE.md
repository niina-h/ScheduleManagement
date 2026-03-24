# 予定管理システム — 初回リリース手順書

## 前提条件

| 項目 | 要件 |
|------|------|
| OS | Windows 10 / 11（64bit） |
| Python | 3.9 以上（3.12 推奨） |
| ネットワーク | 利用者PCから運用PCの 5000 番ポートへアクセス可能であること |
| 権限 | 運用PCの管理者権限（タスクスケジューラ登録時のみ） |

---

## 手順

### 1. Python のインストール（未導入の場合）

```cmd
winget install Python.Python.3.12
```

または https://www.python.org/downloads/ からインストーラーを実行。

> **重要**: インストール時に「Add Python to PATH」にチェックを入れること。

インストール後の確認：

```cmd
python --version
```

---

### 2. アプリケーションの配置

開発PCからプロジェクトフォルダを運用PCへコピーする。

**推奨配置先**: `C:\Apps\ScheduleManagement\`

#### 方法A: robocopy（ネットワーク経由）

```cmd
robocopy "開発元パス" "\\運用PC名\C$\Apps\ScheduleManagement" /E /XD .git .claude __pycache__ /XF web_app.db .secret_key
```

#### 方法B: USB / 共有フォルダ

フォルダを丸ごとコピー。以下は **コピー不要**：

| 除外対象 | 理由 |
|----------|------|
| `.git/` | Git 履歴（運用には不要） |
| `.claude/` | Claude Code 開発設定 |
| `__pycache__/` | Python キャッシュ（自動生成） |
| `db/web_app.db` | 開発用DB（運用先で自動生成） |
| `db/.secret_key` | 開発用鍵（運用先で自動生成） |

---

### 3. 初回セットアップ

運用PC上で `deploy\install.bat` をダブルクリック。

自動で以下が実行される：

1. Python バージョン確認
2. 依存パッケージインストール（Flask, openpyxl, waitress）
3. アプリケーション動作確認

---

### 4. サーバーの自動起動登録

PC再起動後もサーバーが自動起動するよう、タスクスケジューラに登録する。

```
deploy\register_taskscheduler.bat を右クリック →「管理者として実行」
```

自動で以下が実行される：

1. タスク「ScheduleServer」を登録（PC起動時に SYSTEM 権限で実行）
2. サーバーを即座に起動

---

### 5. 動作確認

#### 運用PC上で確認

ブラウザで http://localhost:5000 にアクセス → ログイン画面が表示されること。

#### 他のPCから確認

ブラウザで `http://（運用PCのIPアドレス）:5000` にアクセス。

> 運用PCのIPアドレス確認: `ipconfig` コマンドの IPv4 アドレス

---

### 6. ファイアウォール設定（必要な場合）

他のPCからアクセスできない場合、Windows ファイアウォールでポート 5000 を許可する。

```cmd
netsh advfirewall firewall add rule name="Schedule Server" dir=in action=allow protocol=tcp localport=5000
```

---

### 7. 初期ユーザーの確認

初回起動時に `data/users.json` のサンプルユーザーがDBに自動登録される。

運用に合わせて、マスタ権限のユーザーでログイン後に管理者ダッシュボードからユーザーを追加・編集すること。

---

## 配置後のディレクトリ構成

```
C:\Apps\ScheduleManagement\
├── run_production.py       ← 本番サーバー起動スクリプト
├── requirements_web.txt
├── data/
│   └── users.json          ← 初回移行用（移行後は不要）
├── db/                     ← 自動生成される
│   ├── web_app.db          ← 本番データベース
│   └── .secret_key         ← セッション暗号化キー
├── output/
├── reports/
│   └── tpl/                ← Excel出力テンプレート
├── deploy/
│   ├── install.bat
│   ├── start_server.bat
│   ├── update.bat
│   ├── register_taskscheduler.bat
│   └── register_service.bat
└── web_app/                ← アプリケーション本体
```

---

## 管理コマンド

| 操作 | コマンド |
|------|---------|
| サーバー状態確認 | `schtasks /Query /TN "ScheduleServer"` |
| 手動起動 | `schtasks /Run /TN "ScheduleServer"` |
| 手動停止 | `schtasks /End /TN "ScheduleServer"` |
| タスク登録解除 | `schtasks /Delete /TN "ScheduleServer" /F` |

---

## トラブルシューティング

| 症状 | 対処 |
|------|------|
| ログインページが表示されない | `schtasks /Query /TN "ScheduleServer"` でタスク状態を確認 |
| 他のPCからアクセスできない | ファイアウォール設定（手順6）を確認 |
| ポート5000が使用中 | `deploy\start_server.bat` の `PORT` を変更（例: 8080） |
| pip install 失敗 | プロキシ環境の場合 `pip install --proxy http://proxy:port -r requirements_web.txt` |
