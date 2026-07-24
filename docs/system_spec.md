# ScheduleManagement システム仕様書

- 最終更新：2026-07-24

## 1. システム概要

| 項目 | 内容 |
|------|------|
| 目的 | チームの週間予定・日次実績・ガントチャート進捗を登録・共有・Excel出力するWebアプリ |
| 技術 | Python 3 / Flask / SQLite / Bootstrap 5 / openpyxl |
| 起動 | 開発：`python run_web.py`、本番：`python run_production.py` → http://localhost:5000 |
| TOP画面 | 週間予定画面（ログイン後のデフォルト） |

---

## 2. ユーザーロール体系

### 2-1. 4つのロール

現行は**4段階ロール体系**（一般 → 管理職 → 所属長 → システム管理者）。旧ロール文字列「マスタ」「ユーザー」は`normalize_role()`で自動的に新ロールへ正規化されるため、DBやセッションに旧値が残っていても判定関数は正しく動作する。

| ロール（DB格納値） | 役割 | 概要 |
|--------|------|------|
| **一般** | 一般スタッフ | 自分のデータのみ参照・編集 |
| **管理職** | 部門管理者 | 自部署（直属部下、`manager_id`で判定）のメンバーのみ参照・編集 |
| **所属長** | 複数所属の統括者 | 自分に割り当てられた所属（`user_affiliation`テーブル、複数可）のメンバーのみ参照・編集。他所属には一切関与不可 |
| **システム管理者**（旧「マスタ」） | 最高管理者 | 全所属の参照・編集。所属を切替えて所属長と同じ操作も可能。所属マスタ管理・ユーザー並び替え等のシステム設定も可能 |

### 2-2. 権限判定関数（`web_app/auth_helpers.py`）

| 関数 | 意味 |
|------|------|
| `is_system_admin(role)` | システム管理者（旧マスタ）か |
| `is_dept_head(role)` | 所属長か |
| `is_manager(role)` | 管理職か |
| `is_privileged(role)` | システム管理者・所属長・管理職のいずれか（一般以外なら True） |
| `is_master(role)` | `is_system_admin`への非推奨エイリアス（旧コード互換用） |
| `can_access_user(login_user, target_user)` | 対象ユーザーへのアクセス可否。システム管理者=全員可、所属長=`user_affiliation`所属一致、管理職=`manager_id`一致（未設定時は同一部署へフォールバック） |
| `get_effective_depts()` | 操作対象にできる所属名リストを返す（所属切替・所属長の複数所属対応） |
| `get_active_scope_dept()` | 画面のデータ絞り込みに使う実効所属を返す |
| `can_set_password_for(operator_role, target_role)` | パスワード設定可否（システム管理者は全員可、管理職は一般・管理職のみ） |

### 2-3. 権限マトリックス

| 操作 | 一般 | 管理職 | 所属長 | システム管理者 |
|------|:--------:|:------:|:------:|:------:|
| 自分の週間予定・日次実績・タスクマスター編集 | ○ | ○ | ○ | ○ |
| 他ユーザーの予定・実績閲覧・編集 | × | ○(自部署) | ○(担当所属) | ○(全員) |
| 振り返りコメントレビュー・部下の予定実績振り返り | × | ○(自部署) | ○(担当所属) | ○(全員) |
| ガントチャートの親タスク作成・編集 | × | ○ | ○ | ○ |
| ユーザー追加・削除、所属マスタ管理 | × | × | × | ○ |
| パスワード設定 | × | ○(一般・管理職) | ○(担当所属) | ○(全員) |
| データ移行機能（Excel） | × | × | × | ○ |
| 操作ログ閲覧 | × | ○ | ○ | ○ |

> **既知の不整合**：一部のルート（`tasks.py`の区分管理画面、`project_tasks.py`の旧ガント画面、`planner.py`の一部判定）は、4段階化以前の旧2値ロール文字列（`"マスタ"`/`"管理職"`）と直接比較する判定が残っており、所属長が正しく扱われない箇所がある。改修候補として `docs/進捗管理/改善タスク台帳.md` で追跡する。

初期ユーザーは `data/users.json` に定義し、初回起動時に自動的にデータベースへ取り込まれる。以後は管理者ダッシュボードから追加・削除する。

---

## 3. DBスキーマ（16テーブル）

### users（ユーザーマスタ）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| name / last_name / first_name | TEXT | 氏名（姓名分割対応） |
| login_id | TEXT | ログインID |
| role | TEXT | ロール（4段階） |
| dept | TEXT | 主所属部署名 |
| std_hours_am / std_hours_pm / std_hours | REAL | 基本勤務時間 |
| password_hash | TEXT | パスワードハッシュ |
| remember_token / remember_token_expiry | TEXT | 記憶ログイン用トークン |
| display_order | INTEGER | 表示順 |
| manager_id | INTEGER FK | 上長ユーザーID（管理職判定に使用） |

### user_affiliation（所属長の担当所属）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| user_id | INTEGER FK | 所属長のユーザーID |
| dept_name | TEXT | 担当する所属部署名（複数行で複数所属を表現） |

### task_master（ユーザー別作業マスタ）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| user_id | INTEGER FK | ユーザーID |
| task_name | TEXT | 作業名 |
| display_order | INTEGER | 表示順 |
| default_hours | REAL | デフォルト時間 |
| category_id / subcategory_id | INTEGER FK | 大区分・中区分ID |

### weekly_schedule（週間予定）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| user_id | INTEGER FK | ユーザーID |
| week_start | TEXT | 週開始日（月曜・YYYY-MM-DD） |
| day_of_week | INTEGER | 曜日（0=月〜4=金） |
| time_slot | TEXT | am / pm |
| slot_index | INTEGER | 枠番号（0〜4） |
| task_name / hours / subcategory_name | TEXT/REAL/TEXT | 作業名・時間・中区分名 |
| project_task_id | INTEGER FK | ガントチャート反映元のタスクID（反映していない手動入力はNULL） |
| updated_at / updated_by | TEXT | 更新日時・更新者 |

### daily_result（日次実績）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| user_id / date / time_slot / slot_index | — | 週間予定と同様のキー構造 |
| task_name / hours / subcategory_name | — | 実績内容 |
| is_carryover | INTEGER | 繰越フラグ（0/1） |
| defer_date | TEXT | リスケ先日付 |
| project_task_id | INTEGER FK | 紐づくタスクID（進捗自動連動に使用） |

### daily_comment（日次コメント）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| user_id / date | — | キー |
| reflection | TEXT | 振り返り |
| action | TEXT | 今後の対策・懸念（朝礼での気づき） |
| admin_comment | TEXT | 上長コメント |

### weekly_leave（週間休暇）
| user_id / week_start / day_of_week / leave_type |

### carryover（繰越タスク）
| user_id / task_name / original_date / planned_hours / resolved |

### dept_master（部署マスタ）
| id / dept_name(UNIQUE) / display_order |

### task_category（作業大区分）／ task_subcategory（作業中区分）
| id / name / display_order（中区分は category_id FK を持つ） |

### mail_settings（メール設定）
| カラム | 型 | 概要 |
|--------|-----|------|
| role | TEXT PK | キー値。「管理職」「マスタ」「マスタ_週次管理報告」「管理職_備考」「ユーザー_{user_id}」等。**権限ロールとは別概念**（ユーザー個別キーを含む） |
| to_address / cc_address / bcc_address | TEXT | 宛先 |
| subject_template / body_template | TEXT | テンプレート |

### operation_log（操作ログ）
| id / user_id / user_name / action_type / detail / ip_address / created_at |

### project_task（プロジェクトタスク・ガントチャート・イベント兼用、最も拡張されたテーブル）
| カラム | 型 | 概要 |
|--------|-----|------|
| id | INTEGER PK | 自動採番 |
| category_id / subcategory_id | INTEGER FK | 大区分・中区分 |
| task_name / description | TEXT | タスク名・詳細 |
| start_date / end_date | TEXT | 期間 |
| status | TEXT | 未着手／着手／順調／遅れ／完了／停止 |
| progress / delay_days | INTEGER | 進捗率・遅延日数 |
| display_order | INTEGER | 表示順 |
| created_by / updated_by | — | 作成者・更新者 |
| assigned_to / assigned_to_2 | INTEGER FK | 担当者（子タスク用、最大2名） |
| is_milestone / is_event | INTEGER | マイルストーン／イベントフラグ |
| event_start_time / event_end_time | TEXT | イベント時刻 |
| event_member_ids | TEXT | イベント関係者ID（カンマ区切り、3人目以降） |
| planned_hours | REAL | 予定工数 |
| import_to_schedule_1 / import_to_schedule_2 | INTEGER | 担当者ごとの週間予定反映フラグ |
| parent_task_id | INTEGER FK | 親タスクID（新ガントの親子ツリー構造） |
| member_ids | TEXT | 親タスクのメンバーID（カンマ区切り） |

### company_holiday（会社休日マスタ）
| id / holiday_date(UNIQUE) / holiday_name / created_by |

### routine_schedule（定例作業）
| カラム | 型 | 概要 |
|--------|-----|------|
| id / user_id | — | キー |
| task_name / subcategory_name / default_hours | — | 内容 |
| row_number | INTEGER | 区分行番号（AM=1〜3／PM=6〜8） |
| days | TEXT | 曜日フラグ（例 "1,1,1,1,1"） |
| fill_direction | TEXT | 空きスロットへの詰め方向（top／bottom） |

---

## 4. 画面一覧・URL

### 4-1. 認証（`auth.py`、url_prefixなし）
| URL | メソッド | 概要 |
|-----|---------|------|
| `/login` | GET/POST | ログインID＋4桁PIN。記憶トークンで自動ログイン対応 |
| `/reset_password_and_login` | POST | 初回パスワード設定＋ログイン |
| `/logout` | GET | ログアウト |
| `/switch_dept` | GET | システム管理者・所属長の所属切替 |

### 4-2. 週間予定（`schedule.py`）
| URL | メソッド | 概要 |
|-----|---------|------|
| `/schedule` | GET | 週間予定表示（`?week=YYYY-MM-DD`, `?user_id=N`） |
| `/schedule/save` | POST | 予定保存 |
| `/schedule/copy_last_week` | POST | 先週コピー |
| `/schedule/clear` | POST | 全クリア |
| `/schedule/import_tasks` | POST | タスク管理タスクをAMスロットへインポート |
| `/schedule/import_events` | POST | イベントの自動配置 |
| `/schedule/import_tasks_and_events` | POST | **ガントチャート反映**（定例→タスク→イベントの順で対象週を再構築。当週は禁止・対象週は全消去してから再構築） |

**データ構造**: 月〜金 × AM/PM × 5枠 = 50セル。各セルに作業名・時間を入力。

### 4-3. 日次実績・振り返り（`daily.py`、url_prefixなし）
| URL | メソッド | 概要 |
|-----|---------|------|
| `/daily/today` | GET | 当日へリダイレクト |
| `/daily/<date_str>` | GET | 実績入力画面。予定vs実績の差分（前倒し／予定通り／遅れ）を表示 |
| `/daily/save` | POST | 実績・コメント保存 |
| `/daily/resolve_carryover/<id>` | POST | 繰越を手動解決 |
| `/daily/save_admin_comment` | POST | 上長コメント保存 |
| `/daily/comments-review` | GET | 配下メンバーの振り返りコメントを期間横断で一覧（管理職以上） |
| `/daily/team-progress` | GET | **部下の予定・実績振り返り一覧**（配下メンバー全員の当日差分を遅れの大きい順に表示、管理職以上） |

### 4-4. 作業マスタ（`tasks.py`、url_prefix `/tasks`）
| URL | メソッド | 概要 |
|-----|---------|------|
| `/tasks/` | GET | 作業一覧 |
| `/tasks/add`,`/update-hours/<id>`,`/delete/<id>`,`/move/<id>/<dir>`,`/swap-order` | POST | 作業マスタCRUD・並び替え |
| `/tasks/categories` | GET | 大区分・中区分管理画面 |
| `/tasks/categories/*`,`/subcategories/*` | POST | 区分の追加・削除・並び替え |

### 4-5. タスク管理（旧方式）・定例作業・イベント（`project_tasks.py`、url_prefix `/project-tasks`）
| URL | メソッド | 概要 |
|-----|---------|------|
| `/project-tasks/` | GET | タスク一覧（タスク／定例／イベント タブ） |
| `/project-tasks/routine` | GET | 定例作業専用画面 |
| `/project-tasks/events` | GET | イベント専用画面（月カレンダー＋一覧） |
| `/project-tasks/add` | POST | タスク／イベント追加 |
| `/project-tasks/update/<id>`,`/bulk-update`,`/delete/<id>` | POST | タスク更新・一括更新・削除（管理職以上、`_can_touch_gantt_task`で権限チェック） |
| `/project-tasks/import-brabio` | POST | ブラビオExcelからタスクインポート |
| `/project-tasks/routine/save`,`/routine/delete/<id>` | POST | 定例作業の登録・削除 |
| `/project-tasks/overview` | GET | 全体俯瞰ダッシュボード |
| `/project-tasks/dashboard`,`/dashboard/api` | GET | 個人別進捗ダッシュボード |
| `/project-tasks/gantt` | GET | **旧ガントチャート画面**（大区分/中区分グルーピング表示） |
| `/project-tasks/gantt/update-dates/<id>`,`/gantt/update-fields/<id>`,`/gantt/reorder` | POST | 旧ガントのドラッグ操作API（管理職以上） |
| `/project-tasks/gantt/export` | GET | 旧ガントチャートExcel出力 |
| `/project-tasks/export-migration` | GET | **移行用フラットExcel出力**（システム管理者のみ、planner.pyへのデータ移行用） |

### 4-6. ガントチャート（新方式・視覚入力）（`planner.py`、url_prefix `/planner`）
| URL | メソッド | 概要 |
|-----|---------|------|
| `/planner/` | GET | **ガント入力画面**。タイムライン上でバーをドラッグして期間を描画し、親子ツリー構造でタスクを登録する |
| `/planner/export`,`/export-pdf` | GET | 表示中ツリー・期間と一致するExcel/PDF出力 |
| `/planner/save` | POST | 視覚入力の保存（JSON差分：新規/更新/削除を`project_task`へ反映。親作成は管理職以上のみ） |
| `/planner/import-excel` | POST | **移行用フラットExcelの取込**（システム管理者のみ。id列一致で更新／空欄で新規、`parent_task_id`列で親子関係を2パスで解決） |

### 4-7. 管理ダッシュボード（`admin.py`、url_prefix `/admin`）
| URL | メソッド | 概要 |
|-----|---------|------|
| `/admin/` | GET | 週間予定登録状況・当日実績状況・ユーザー一覧 |
| `/admin/users/add`,`/delete/<id>`,`/update_hours/<id>`,`/set_password/<id>`,`/clear_password/<id>`,`/bulk_update` | POST | ユーザー管理 |
| `/admin/depts/add`,`/delete/<id>`,`/update/<id>` | POST | 部署マスタ管理 |
| `/admin/users/reorder` | POST | ユーザー並び替え |
| `/admin/db-download` | GET | SQLite DBファイルダウンロード（システム管理者のみ） |
| `/admin/api/daily_status` | GET | 当日実績入力状況API（ポーリング用） |
| `/admin/logs` | GET | 操作ログ一覧 |
| `/admin/master/export/<table_key>`,`/master/export-all` | GET | マスタデータCSVエクスポート（9テーブル対応） |
| `/admin/master/import` | POST | マスタデータCSVインポート（UPSERT対応） |
| `/admin/company-holiday/add`,`/delete/<id>` | POST | 会社休日管理 |

### 4-8. Excel出力・日次業務報告（`export.py`、url_prefix `/export`）
| URL | メソッド | 概要 |
|-----|---------|------|
| `/export/daily/<date_str>` | GET | 自分の日報Excel |
| `/export/admin_daily/<date_str>` | GET | 全員日報Excel（シート別） |
| `/export/admin_report/<date_str>` | GET | 管理者用日次報告（自分＋部下サマリ＋翌日予定） |
| `/export/my` | GET | 自分の週間予定Excel |
| `/export/team_week` | GET | 全メンバー週間予定Excel |
| `/export/user/<id>` | GET | 指定ユーザーの週間予定実績表Excel |
| `/export/multi_week` | GET | 自分の前週/今週/来週3シートExcel |
| `/export/report/download`,`/report/print`,`/report/team` | GET | 日報の追加出力形式 |
| `/export/import` | GET/POST | 週間予定Excelインポート（システム管理者のみ） |

### 4-9. 日報メール（`mail_report.py`、url_prefix `/mail-report`）
| URL | メソッド | 概要 |
|-----|---------|------|
| `/mail-report/preview` | GET | 管理職・システム管理者用メールプレビュー（業務内容サマリー＋開発状況を含む動的生成本文） |
| `/mail-report/download_eml` | GET | メール内容の.emlダウンロード |
| `/mail-report/print-master` | GET | システム管理者用メール印刷ページ |
| `/mail-report/save-address`,`/save-friday-report`,`/save-mgr-remarks` | POST | 宛先・金曜報告文・備考の保存 |
| `/mail-report/settings` | GET/POST | メール設定画面（システム管理者のみ） |
| `/mail-report/user-preview` | GET | 一般ユーザー用日報メールプレビュー |
| `/mail-report/save-user-address`,`/save-user-body` | POST | ユーザー用宛先・本文保存 |
| `/mail-report/download-user-eml` | GET/POST | ユーザー用メール.emlダウンロード |

### 4-10. ヘルプ（`help.py`）・API（`api.py`）
| URL | メソッド | 概要 |
|-----|---------|------|
| `/help`,`/help/<page>` | GET | ヘルプ・手順書 |
| `/api/today-events` | GET | 本日のイベント一覧JSON（ブラウザ通知用） |

---

## 5. 主要な業務フロー

### 5-1. ガントチャート起点の全体像（目指す姿）

```
① ガントチャートで先々の計画を立てる
   （新方式：/planner/ で視覚的にバーを描く → タスク一覧化）
      ↓
② 翌週の週間予定を立てる
   ・「ガントチャート反映」ボタン（/schedule/import_tasks_and_events）で
     該当週の自分の予定にタスクを反映
   ・定例作業も同時に登録・割当（詰め方向を選択可能）
   ・反映されたタスクに予定時間を手入力して保存 → 週間予定作成 完了
      ↓
③ 実績の入力（/daily/<date_str>）
   ・当日、各ユーザーが予定に対して実績を入力して保存
      ↓
④ 振り返り
   ・個人：daily.html でその日のガント予定と実績の差分を表示
   ・管理職・所属長・システム管理者：/daily/team-progress で
     部下全員の予定実績差分を一覧
```

### 5-2. ガントチャートでのタスク登録（新旧2系統が並存）

- **新方式（`planner.py`）**：`/planner/` でタイムライン上にバーをドラッグ描画 → `POST /planner/save` でJSON差分（親子ツリー構造の行）を送信 → `add_project_task`/`update_project_task`/`set_project_task_parent`/`set_project_task_members` で `project_task` テーブルへ反映。親（レベル0）タスクの作成・編集は管理職以上限定。
- **旧方式（`project_tasks.py`）**：`/project-tasks/` のフォームで直接入力し `/project-tasks/add` や `/update/<id>` へPOST。旧ガントチャート画面（`/project-tasks/gantt`）ではドラッグによる日付変更のみ可能。
- **移行経路**：`/project-tasks/export-migration` で全タスクをフラットExcel出力 → `/planner/import-excel` で新ガント（親子ツリー）構造として取込。`id`列一致で更新・空欄で新規、`parent_task_id`列で親子関係を2パスで解決（循環参照は検知してスキップ）。

### 5-3. 週間予定へのガント反映

`POST /schedule/import_tasks_and_events` が中核ロジック。当週への反映は誤操作防止のため禁止（前週・次週以降のみ）。処理順：

1. 対象週の `weekly_schedule` を一旦全クリア
2. `apply_routine_to_week()` で**定例作業**を区分（AM=row1-3／PM=row6-8）内の空きスロットへ、定例ごとに設定した `fill_direction`（top／bottom）に従って詰める
3. `import_tasks_to_weekly_schedule()` で `project_task` のうち、担当者一致・期間重複・`import_to_schedule_1/2` フラグONのものをAM/PM交互配分で配置
4. `import_events_to_weekly_schedule()` でイベント（`is_event=1`）を開始時刻からAM/PM判定し配置

個別に「タスクのみ反映」「イベントのみ反映」も利用可能。

### 5-4. 日次実績の入力

`GET /daily/<date_str>` で当日の週間予定・実績・繰越・イベント・タスク進捗を統合表示 → `POST /daily/save` で保存。リスケ（翌日以降の`weekly_schedule`へ転記）、繰越（`carryover`テーブル更新）、`project_task_id`紐づけ実績は `sync_daily_progress_to_task()` でタスクの `progress` に自動反映される。

### 5-5. 予定vs実績の振り返り

- **個人向け**（`daily.html`）：`get_task_plan_vs_actual()` で当日時点の担当タスクごとの予定進捗率・実績進捗率を比較し、遅れ／前倒しを表示。
- **管理職・所属長・システム管理者向け**（`team_progress.html`）：`get_daily_plan_vs_actual_for_users()` で配下メンバー全員の当日予定時間・実績時間・差分をまとめて一覧表示（遅れが大きい順にソート、グラフなしで数値・バッジ中心のシンプルなUI）。
- 関連機能として `/daily/comments-review` で配下メンバーの振り返りコメントを期間横断でレビューし、上長コメント未入力を強調表示。

### 5-6. メール日報（宛先ごとに本文生成方式が異なる）

- **システム管理者用**（`_build_master_body`）：動的生成。業務内容セクションは状態別（遅れ→着手→未着手→順調→完了→停止）のサマリー行付き一覧。開発状況セクションは大区分「開発」の project_task から担当者・進捗・状態を表示。メンバー別AM/PM実績サマリ、計画/突発/リスケの時間比率、金曜日は「管理業務のご報告」セクションを追加挿入。
- **管理職自身用**（`_build_mgr_self_body`）：固定フォーマットテンプレート。
- **一般ユーザー用**（`user_preview`）：`mail_settings.role = "ユーザー_{user_id}"` という個別キーで管理。本文はデフォルトテンプレートをユーザーが自由編集。自動集計は行わない。

いずれも `.eml` ファイル生成、または `mailto:` リンクでOutlook等の既定メールソフトを起動する。

### 5-7. Excel移行機能

- `project_tasks.py` の `export_migration()`（システム管理者専用）が全 `project_task` をフラット形式（id, parent_task_id, task_name, 担当者1/2, 開始日, 終了日, 状態, 進捗%, メンバー）でExcel出力。
- `planner.py` の `import_excel()`（システム管理者専用）が `import_migration_excel()` を呼び、id列一致で更新・空欄で新規追加、`parent_task_id`列で親子ツリーを2パス解決して反映。
- 別系統として `import_brabio()` は外部「ブラビオ」Excel形式からのタスクインポート（日常運用用、移行専用とは別機能）。

---

## 6. 画面遷移概要

```
ログイン画面
    ↓ ログイン成功
┌──────────────────────────── ナビゲーションバー ────────────────────────────┐
│ 週間予定 │ 本日実績 │ タスク管理 │ ガントチャート │ 管理(※) │ ヘルプ │ ログアウト │
└──────────┴──────────┴────────────┴────────────────┴──────────┴────────┴───────────┘
                                                       ※管理職以上のみ表示

週間予定 (schedule.html)                日次実績 (daily.html)
 ・月〜金×AM/PM×5枠の入力グリッド        ・予定を自動表示→実績を上書き入力
 ・先週コピー／クリア／ガント反映         ・振り返り・予定実績差分バッジ
 ・当週はガント反映不可                   ・リスケ／繰越ボタン

タスク管理・ガントチャート
 ・/planner/     新方式：視覚的にバーを描いてタスク登録
 ・/project-tasks/gantt  旧方式：大区分/中区分グルーピング表示
 ・/project-tasks/       タスク一覧・定例作業・イベント（タブ切替）

管理ダッシュボード (admin.html)
 ・週間予定登録状況・当日実績状況の一覧
 ・ユーザー・所属・会社休日の管理
 ・操作ログ／Excel出力／メール設定／マスタCSVインポート

振り返り（管理職以上）
 ・/daily/comments-review   配下の振り返りコメント期間横断レビュー
 ・/daily/team-progress     配下全員の予定実績差分一覧
```

---

## 7. 作業区分の階層構造

```
task_category（大区分）
  └─ task_subcategory（中区分）
       └─ task_master（作業）← ユーザーごとに設定

例:
  大区分: 開発
    中区分: AI開発
      作業: 図面解析対応, 学習データ作成
    中区分: システム開発
      作業: 要件定義確認
  大区分: 管理
    中区分: 定例作業
      作業: 朝礼, 週次MTG
```

---

## 8. 主要テンプレート一覧

| ファイル | 画面概要 |
|---|---|
| `base.html` | 全画面共通レイアウト（ナビバー） |
| `login.html` | ログイン画面 |
| `schedule.html` | 週間予定入力・表示画面 |
| `daily.html` | 日次実績入力・振り返り画面（個人） |
| `daily_report_print.html` | 日次業務報告の印刷用ページ |
| `admin.html` | 管理者ダッシュボード |
| `admin_logs.html` | 操作ログ一覧画面 |
| `comments_review.html` | 配下メンバーの振り返りコメント期間横断レビュー画面 |
| `team_progress.html` | 部下の予定・実績振り返り一覧画面 |
| `tasks.html` | 作業マスタ管理画面 |
| `tasks/categories.html` | 作業大区分・中区分管理画面 |
| `project_tasks.html` | タスク管理・定例作業・イベント画面（タブ切替共通） |
| `project_tasks_gantt.html` | 旧ガントチャート画面（大区分/中区分グルーピング） |
| `project_tasks_dashboard.html` | 個人別進捗ダッシュボード |
| `project_tasks_overview.html` | 全体ステータス俯瞰ダッシュボード |
| `gantt_input_test.html` | 新ガントチャート（視覚入力）画面 |
| `import_schedule.html` | 週間予定Excelインポート画面 |
| `mail_report_preview.html` | 管理職・システム管理者用日報メールプレビュー |
| `mail_report_print.html` | システム管理者用メール本文印刷ページ |
| `mail_report_settings.html` | メール設定画面 |
| `mail_user_preview.html` | 一般ユーザー用日報メールプレビュー |
| `help.html` | ヘルプ・操作手順ページ |

---

## 9. 既知の技術的負債・改修候補

- **旧2値ロール文字列への依存が一部に残存**：`tasks.py`の区分管理画面（`_require_privileged`）、`project_tasks.py`の`gantt()`、`planner.py`の一部が `role == "マスタ"` / `"管理職"` の直接比較を行っており、所属長ロールがこれらの箇所で正しく扱われない可能性がある。
- **新旧ガントチャートの並存**：`planner.py`（新方式）と`project_tasks.py`の`gantt`関連ルート（旧方式）が並存している。当面は両方維持し、リリース安定後に旧方式の廃止を検討する方針（詳細は `docs/進捗管理/改善タスク台帳.md` 参照）。

進行中のタスク・修正履歴の詳細は [`docs/進捗管理/改善タスク台帳.md`](進捗管理/改善タスク台帳.md) を参照。
