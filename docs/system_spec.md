# ScheduleManagement システム仕様書

## 1. システム概要

| 項目 | 内容 |
|------|------|
| 目的 | チームの週間予定・日次実績を登録・共有・Excel出力するWebアプリ |
| 技術 | Python 3 / Flask / SQLite (WAL) / Bootstrap 5 / openpyxl |
| 起動 | `python run_web.py` → http://localhost:5000 |
| TOP画面 | 週間予定画面（ログイン後のデフォルト） |

---

## 2. ユーザーロール体系

### 2-1. 3つのロール

| ロール | 役割 | 概要 |
|--------|------|------|
| **マスタ** | 最高管理者 | 全ユーザーのデータ参照・編集。部署・区分・メール設定管理。ユーザー追加削除。 |
| **管理職** | 部門管理者 | 自部署メンバーのデータ参照・編集。上長コメント入力。パスワード設定。 |
| **ユーザー** | 一般スタッフ | 自分のデータのみ参照・編集。パスワード不要。 |

### 2-2. 権限マトリックス

| 操作 | ユーザー | 管理職 | マスタ |
|------|:--------:|:------:|:------:|
| 自分の週間予定編集 | ○ | ○ | ○ |
| 自分の日次実績入力 | ○ | ○ | ○ |
| 自分の作業マスタ管理 | ○ | ○ | ○ |
| 他ユーザーの予定閲覧・編集 | × | ○(自部署) | ○(全員) |
| 他ユーザーの実績閲覧 | × | ○(自部署) | ○(全員) |
| 上長コメント入力 | × | ○(自部署) | ○(全員) |
| ユーザー追加・削除 | × | × | ○ |
| ユーザー一括編集 | × | ○(自部署) | ○(全員) |
| パスワード設定 | × | ○(ユーザー・管理職) | ○(全員) |
| 部署マスタ管理 | × | × | ○ |
| 大区分・中区分管理 | × | ○ | ○ |
| メール設定 | × | ○(自部署) | ○(全員) |
| Excel出力 | × | ○(自部署) | ○(全員) |
| 操作ログ閲覧 | × | ○ | ○ |

---

## 3. DBスキーマ（12テーブル）

### users（ユーザーマスタ）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| name | TEXT UNIQUE | ユーザー名 |
| role | TEXT | ロール（マスタ/管理職/ユーザー） |
| dept | TEXT | 部署名 |
| std_hours_am | REAL | AM基本時間 |
| std_hours_pm | REAL | PM基本時間 |
| std_hours | REAL | 合計基本時間 |
| password_hash | TEXT | パスワードハッシュ |
| remember_token | TEXT | 記憶ログイン用トークン |
| remember_token_expiry | TEXT | トークン有効期限 |
| display_order | INTEGER | 表示順 |
| manager_id | INTEGER FK | 上長ユーザーID |

### task_master（ユーザー別作業マスタ）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| user_id | INTEGER FK | ユーザーID |
| task_name | TEXT | 作業名 |
| display_order | INTEGER | 表示順 |
| default_hours | REAL | デフォルト時間 |
| category_id | INTEGER FK | 大区分ID |
| subcategory_id | INTEGER FK | 中区分ID |

### weekly_schedule（週間予定）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| user_id | INTEGER FK | ユーザーID |
| week_start | TEXT | 週開始日（月曜・YYYY-MM-DD） |
| day_of_week | INTEGER | 曜日（0=月〜4=金） |
| time_slot | TEXT | am / pm |
| slot_index | INTEGER | 枠番号（0〜4） |
| task_name | TEXT | 作業名 |
| hours | REAL | 時間 |
| subcategory_name | TEXT | 中区分名 |
| created_at | TEXT | 作成日時 |
| updated_at | TEXT | 更新日時 |
| updated_by | TEXT | 更新者名 |

### daily_result（日次実績）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| user_id | INTEGER FK | ユーザーID |
| date | TEXT | 日付（YYYY-MM-DD） |
| time_slot | TEXT | am / pm |
| slot_index | INTEGER | 枠番号（0〜4） |
| task_name | TEXT | 作業名 |
| hours | REAL | 時間 |
| subcategory_name | TEXT | 中区分名 |
| is_carryover | INTEGER | 繰越フラグ（0/1） |
| defer_date | TEXT | 遅延元日付 |
| updated_at | TEXT | 更新日時 |
| updated_by | TEXT | 更新者名 |

### daily_comment（日次コメント）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| user_id | INTEGER FK | ユーザーID |
| date | TEXT | 日付 |
| reflection | TEXT | 振り返り |
| action | TEXT | 朝礼での気づき |
| admin_comment | TEXT | 上長コメント |
| updated_at | TEXT | 更新日時 |
| updated_by | TEXT | 更新者名 |

### weekly_leave（週間休暇）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| user_id | INTEGER FK | ユーザーID |
| week_start | TEXT | 週開始日 |
| day_of_week | INTEGER | 曜日 |
| leave_type | TEXT | 休暇種別 |

### carryover（繰越タスク）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| user_id | INTEGER FK | ユーザーID |
| task_name | TEXT | 作業名 |
| original_date | TEXT | 元の日付 |
| planned_hours | REAL | 予定時間 |
| resolved | INTEGER | 完了フラグ |

### dept_master（部署マスタ）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| dept_name | TEXT UNIQUE | 部署名 |
| display_order | INTEGER | 表示順 |

### task_category（作業大区分）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| name | TEXT UNIQUE | 大区分名 |
| display_order | INTEGER | 表示順 |

### task_subcategory（作業中区分）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| category_id | INTEGER FK | 大区分ID |
| name | TEXT | 中区分名（大区分内でUNIQUE） |
| display_order | INTEGER | 表示順 |

### mail_settings（メール設定）
| カラム | 型 | 概要 |
|--------|-----|------|
| role | TEXT PK | ロール名（管理職/マスタ） |
| to_address | TEXT | TO |
| cc_address | TEXT | CC |
| bcc_address | TEXT | BCC |
| subject_template | TEXT | 件名テンプレート |
| body_template | TEXT | 本文テンプレート |

### operation_log（操作ログ）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| user_id | INTEGER | ユーザーID |
| user_name | TEXT | ユーザー名 |
| action_type | TEXT | 操作種別 |
| detail | TEXT | 詳細 |
| ip_address | TEXT | IPアドレス |
| created_at | TEXT | 日時 |

---

## 4. 画面一覧・URL

### 4-1. 認証
| URL | メソッド | 画面 | 概要 |
|-----|---------|------|------|
| `/login` | GET/POST | login.html | ユーザー選択・パスワード入力 |
| `/logout` | GET | — | ログアウト→ログイン画面へ |

### 4-2. 週間予定
| URL | メソッド | 画面 | 概要 |
|-----|---------|------|------|
| `/schedule` | GET | schedule.html | 週間予定表示（?week=YYYY-MM-DD, ?user_id=N） |
| `/schedule/save` | POST | — | 予定保存 |
| `/schedule/copy_last_week` | POST | — | 先週コピー |
| `/schedule/clear` | POST | — | 全クリア |

**データ構造**: 月〜金 × AM/PM × 5枠 = 50セル。各セルに作業名・時間を入力。

### 4-3. 日次実績
| URL | メソッド | 画面 | 概要 |
|-----|---------|------|------|
| `/daily/today` | GET | — | 当日へリダイレクト |
| `/daily/<date_str>` | GET | daily.html | 実績入力画面 |
| `/daily/save` | POST | — | 実績・コメント保存 |
| `/daily/<date>/defer` | POST | — | 遅延→翌営業日予定へ |
| `/daily/<date>/carryover` | POST | — | 繰越→翌営業日予定へ |

**データ構造**: AM/PM × 5枠。予定値を自動表示し、実績を上書き入力。振り返り・要点・上長コメント。

### 4-4. 作業マスタ
| URL | メソッド | 画面 | 概要 |
|-----|---------|------|------|
| `/tasks/` | GET | tasks.html | 作業一覧 |
| `/tasks/add` | POST | — | 作業追加 |
| `/tasks/delete/<id>` | POST | — | 作業削除 |
| `/tasks/move/<id>/<dir>` | POST | — | 上下移動 |
| `/tasks/swap-order` | POST | — | 入れ替え（AJAX） |
| `/tasks/categories` | GET | categories.html | 大区分・中区分管理 |

### 4-5. 管理ダッシュボード
| URL | メソッド | 画面 | 概要 |
|-----|---------|------|------|
| `/admin/` | GET | admin.html | ダッシュボード |
| `/admin/users/add` | POST | — | ユーザー追加 |
| `/admin/users/delete/<id>` | POST | — | ユーザー削除 |
| `/admin/users/bulk_update` | POST | — | 一括更新 |
| `/admin/users/reorder` | POST | — | 順序変更（AJAX） |
| `/admin/depts/*` | POST | — | 部署管理 |
| `/admin/assignments/save` | POST | — | 上長割当 |
| `/admin/api/daily_status` | GET | — | 実績状況API（30秒ポーリング） |
| `/admin/logs` | GET | admin_logs.html | 操作ログ |

### 4-6. Excel出力・日次業務報告
| URL | メソッド | 画面 | 概要 |
|-----|---------|------|------|
| `/export/schedule` | GET | — | 週間予定Excel |
| `/export/schedule_with_results` | GET | — | 予定+実績Excel |
| `/export/import` | GET/POST | import_schedule.html | Excelインポート（マスタのみ） |
| `/export/report/download` | GET | — | 日次業務報告Excel（テンプレート形式） |
| `/export/report/print` | GET | daily_report_print.html | 日報印刷用HTML表示（ブラウザ印刷でPDF保存） |
| `/export/report/team` | GET | — | チーム日次業務報告Excel（メンバー別シート） |

**日次業務報告の出力形式**:
- **Excel**: テンプレートファイル（`reports/tpl/日次業務報告_テンプレート.xlsx`）を元に生成。ファイル名は `日次業務報告_氏名_日付.xlsx`。区分列なし・1ページ印刷固定・翌日予定カンマ区切り
- **PDF（ブラウザ印刷）**: Excelテンプレート準拠のHTMLを表示し、ブラウザの印刷機能でPDF保存。外部ライブラリ不要
- **チームExcel**: 管理職・マスタ用。担当メンバー分をシート別にまとめて出力。ファイル名は `日次業務報告_チーム_日付.xlsx`

### 4-7. 日報メール
| URL | メソッド | 画面 | 概要 |
|-----|---------|------|------|
| `/mail-report/preview` | GET | mail_report_preview.html | 管理職・マスタ用メールプレビュー |
| `/mail-report/save-address` | POST | — | 管理職・マスタ用宛先保存 |
| `/mail-report/save-friday-report` | POST | — | 金曜日管理業務報告テキスト保存 |
| `/mail-report/save-mgr-remarks` | POST | — | マスタ備考欄保存（印刷のみ表示） |
| `/mail-report/download_eml` | GET | — | EMLダウンロード |
| `/mail-report/user-preview` | GET | mail_user_preview.html | ユーザー用日報メールプレビュー |
| `/mail-report/save-user-address` | POST | — | ユーザー用宛先保存 |
| `/mail-report/save-user-body` | POST | — | ユーザー用本文テンプレート保存 |
| `/mail-report/download-user-eml` | GET | — | ユーザー用EMLダウンロード |
| `/mail-report/settings` | GET/POST | mail_report_settings.html | メール設定（マスタのみ） |

### 4-8. ヘルプ
| URL | メソッド | 画面 | 概要 |
|-----|---------|------|------|
| `/help` | GET | help.html | ヘルプトップ |
| `/help/<page>` | GET | help.html | 各ページヘルプ |

### 4-9. タスク管理・ガントチャート
| URL | メソッド | 画面 | 概要 |
|-----|---------|------|------|
| `/project-tasks/` | GET | project_tasks.html | タスク一覧（タスク/定例/イベント タブ） |
| `/project-tasks/add` | POST | — | タスク追加（イベントは全ユーザー可） |
| `/project-tasks/update/<id>` | POST | — | タスク更新 |
| `/project-tasks/bulk-update` | POST | — | タスク一括更新 |
| `/project-tasks/delete/<id>` | POST | — | タスク削除 |
| `/project-tasks/gantt` | GET | project_tasks_gantt.html | ガントチャート（ヘッダー固定・イベント行・今日赤枠） |
| `/project-tasks/gantt/export` | GET | — | ガントチャートExcel出力（凡例・印刷設定付き） |
| `/project-tasks/gantt/update-dates/<id>` | POST | — | ガントバードラッグ日程変更 |
| `/project-tasks/gantt/update-fields/<id>` | POST | — | ガントインライン編集（状態・進捗） |
| `/project-tasks/gantt/reorder` | POST | — | ガント行並べ替え |
| `/project-tasks/overview` | GET | project_tasks_overview.html | 全体ステータスダッシュボード |
| `/project-tasks/dashboard` | GET | project_tasks_dashboard.html | 進捗ダッシュボード |

---

## 5. 画面遷移図

```
ログイン画面
    ↓ ログイン成功
    ↓
┌──────────────────────────── ナビゲーションバー ────────────────────────────┐
│  週間予定 │ 本日実績 │ 作業登録 │ 管理(※) │ ヘルプ │ ログアウト          │
└───────────┴──────────┴──────────┴──────────┴────────┴─────────────────────┘
    ↓                                         ※管理職・マスタのみ表示
    ↓
┌── 週間予定 (schedule.html) ─── デフォルト画面 ──┐
│  ・月〜金 × AM/PM × 5枠の入力グリッド          │
│  ・前週/次週ナビ                                │
│  ・先週コピー / クリア                           │
│  ・管理者は他ユーザー選択で閲覧・編集           │
│  ・休暇設定（全休/AM休/PM休）                   │
└─────────────────────────────────────────────────┘

┌── 日次実績 (daily.html) ────────────────────────┐
│  ・当日の予定値を自動表示 → 実績を上書き入力   │
│  ・AM/PM × 5枠                                  │
│  ・振り返り / 朝礼での気づき テキスト入力       │
│  ・上長コメント（管理職・マスタのみ入力可）     │
│  ・遅延 / 繰越ボタン → 翌営業日予定へ移行      │
│  ・前日/翌日ナビ                                │
│  → 管理職日報メール画面へのリンク               │
└─────────────────────────────────────────────────┘

┌── 作業登録 (tasks.html) ────────────────────────┐
│  ・自分の作業一覧（追加/削除/並べ替え）         │
│  ・大区分・中区分の紐付け                        │
│  ・デフォルト時間設定                             │
│  → 大区分・中区分管理（管理職・マスタ）          │
└─────────────────────────────────────────────────┘

┌── 管理ダッシュボード (admin.html) ──────────────┐
│  ・週間予定登録状況（全ユーザー一覧）           │
│  ・本日の実績入力状況（30秒自動更新）           │
│  ・ユーザー設定（追加/削除/パスワード/順序）    │
│  ・部署管理                                      │
│  ・メンバー上長割当                              │
│  → 操作ログ / Excel出力 / メール設定 / インポート│
└─────────────────────────────────────────────────┘
```

---

## 6. 主要データフロー

### 6-1. 週間予定 → 日次実績の連携
```
週間予定（plan）                日次実績（result）
┌─────────────────┐           ┌─────────────────┐
│ user_id          │           │ user_id          │
│ week_start       │           │ date             │
│ day_of_week(0-4) │ ──copy──→ │ time_slot        │
│ time_slot(am/pm) │  初期値   │ slot_index       │
│ slot_index(0-4)  │           │ task_name        │
│ task_name        │           │ hours            │
│ hours            │           │ is_carryover     │
└─────────────────┘           └─────────────────┘
```
- 日次実績画面を開くと、その日の週間予定が初期値として表示される
- ユーザーが上書き入力すると実績として保存される

### 6-2. 遅延・繰越フロー
```
当日実績で未完了 → [遅延/繰越] → 翌営業日の週間予定に自動追加
                                  → carryover テーブルに記録
                                  → daily_result に is_carryover=1 / defer_date 設定
```

### 6-3. Excel出力フロー
```
週間予定 + 日次実績 → openpyxl → テンプレート形式Excel → ダウンロード
  ・計画値と実績値を並記
  ・突発（計画にない作業）にマーク付け
  ・AM/PM ごとのセット比較で判定
```

### 6-3b. 日次業務報告PDF出力フロー（ブラウザ印刷方式）
```
日次実績 + コメント → Jinja2 → 印刷用HTML(daily_report_print.html)
    → ブラウザ新タブで表示 → ユーザーが「印刷/PDF保存」ボタン押下
    → ブラウザの印刷ダイアログで「PDFに保存」を選択
  ・Excelテンプレート準拠のデザイン（紺ヘッダー・背景色・罫線）
  ・外部ライブラリ不要（サーバー側はHTMLを返すのみ）
  ・@media print で操作ボタンを非表示
```

### 6-4. 日報メールフロー
```
管理職用: 自分の実績 + 振り返り + 翌日予定 → メール本文生成
マスタ用: 全メンバー実績 + 大区分別集計 + 計画/突発/リスケ率 → メール本文生成
    → クリップボードに書式付きコピー + mailto: でOutlook起動
```

---

## 7. 作業区分の階層構造

```
task_category（大区分）
  └─ task_subcategory（中区分）
       └─ task_master（作業）← ユーザーごとに設定

例:
  大区分: 開発業務
    中区分: AI開発
      作業: ChatBot実装, プロンプト設計
    中区分: Web開発
      作業: フロントエンド改修
  大区分: 管理業務
    中区分: 定例作業
      作業: 朝礼, 週次MTG
```

---

## 8. ファイル構成

```
ScheduleManagement/
├── run_web.py                    # 起動スクリプト
├── requirements_web.txt          # 依存パッケージ
├── CLAUDE.md                     # Claude Code設定
├── data/users.json               # 初期ユーザーデータ
├── db/
│   ├── web_app.db                # SQLite DB（自動生成）
│   └── .secret_key               # Flask SECRET_KEY
├── output/                       # Excel出力先
├── docs/                         # ドキュメント
└── web_app/
    ├── app.py                    # Flaskアプリファクトリ
    ├── config.py                 # 設定
    ├── database.py               # DB初期化・スキーマ
    ├── models.py                 # データアクセス関数（約1750行）
    ├── auth_helpers.py           # 権限チェック
    ├── log_service.py            # 操作ログ記録
    ├── routes/
    │   ├── auth.py               # 認証
    │   ├── schedule.py           # 週間予定
    │   ├── daily.py              # 日次実績
    │   ├── tasks.py              # 作業マスタ
    │   ├── admin.py              # 管理ダッシュボード
    │   ├── export.py             # Excel出力・インポート
    │   ├── mail_report.py        # 日報メール
    │   └── help.py               # ヘルプ
    ├── templates/                # Jinja2テンプレート（13ファイル）
    └── static/style.css          # カスタムCSS
```

---

## 9. 現在の時間管理の粒度

| レベル | 単位 | 現状 |
|--------|------|------|
| **年間** | — | なし |
| **月間** | — | **なし（今後追加予定）** |
| **週間** | 週（月〜金） | weekly_schedule テーブルで管理 |
| **日次** | 日 | daily_result テーブルで管理 |
| **スロット** | AM/PM × 5枠 | 最小入力単位 |

→ **月計画** を追加することで、月→週→日 の3階層計画が可能になる。
