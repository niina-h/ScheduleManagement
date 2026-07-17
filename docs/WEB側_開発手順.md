# WEB側 開発手順

Python バックエンド + React フロントを **1 システムで同時開発・運用** するための手順書。

> 関連ドキュメント：
> - `共通UI・レイアウト設計ガイド.md`（フロント設計の詳細）
> - `resources/web_template/CLAUDE.md`（フロント側 AI 行動規約）
> - ルートの `CLAUDE.md`（Python 側 AI 行動規約）

---

## 1. 全体像

```
[ ブラウザ ]
     │ HTTP
     ▼
[ React フロント（Vite dev: :5173） ]
     │ apiFetch (/api/...)
     ▼
[ Python FastAPI（uvicorn: :8000） ]
     │
     ▼
[ 各 feature の service / repository ]
```

- フロントは `VITE_API_BASE_URL`（既定 `http://localhost:8000/api`）にアクセス
- Python 側は `python main.py web` で FastAPI を起動
- 同じツール（feature）が Python と React の両方に**対称的に存在**する

---

## 2. 初回セットアップ

### 2.1 Python 側
```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2.2 Web 側
```bash
cd resources/web_template
cp .env.example .env.local         # 必要に応じて値を編集
npm install
```

---

## 3. 開発時の起動順（ターミナル 2 つ）

### ターミナル A：Python API サーバ
```bash
python main.py web --reload
```
→ `http://localhost:8000/docs` で OpenAPI（Swagger UI）を確認できる

### ターミナル B：Vite 開発サーバ
```bash
cd resources/web_template
npm run dev
```
→ `http://localhost:5173` でフロントが立ち上がる

ログイン：`demo` / `demo` （雛形の認証情報。本実装で必ず差し替える）

---

## 4. ツール（feature）追加の対称ペア手順

「集計ツール（`aggregate_tool`）」を追加する例。**Python と Web を必ず対で**追加する。

### 4.1 Python 側
1. `src/features/aggregate_tool/` を作成
2. `controller.py`（CLI 用 `run(args)`）、`service.py`、`model.py` を配置
3. `api.py` を作成し、`router = APIRouter()` + エンドポイント定義
4. `main.py` の `_FEATURES` 辞書に `"aggregate": ("src.features.aggregate_tool", "集計ツール")` を追加
5. `src/features/web/app.py` の `create_app()` に `app.include_router(aggregate_api.router, ...)` を追加
6. `tests/features/aggregate_tool/test_service.py` でユニットテストを書く

### 4.2 Web 側
1. `resources/web_template/src/features/aggregate_tool/` を作成
2. `api.ts`（Python の `/api/aggregate` を呼ぶ）を配置
3. `pages/AggregatePage.tsx` を作成
4. `src/router.tsx` に `<Route path="/aggregate" element={<AuthGuard><AggregatePage /></AuthGuard>} />` を追加
5. `pages/AggregatePage.tsx` 内に `<NavigatableShell breadcrumbs={["ホーム", "集計"]}>` を仕込む

これだけで **CLI（`python main.py aggregate`） / Web 画面（`/aggregate`） / Web API（`/api/aggregate`）が全部動く**。

---

## 5. 認証フロー

```
1. ユーザーがログイン画面（/login）に入力
2. フロント → POST /api/login → Python が token を返す
3. setToken(token) で sessionStorage に保存
4. 以降の apiFetch が Authorization: Bearer <token> を自動付与
5. AuthGuard が保護ルートで isAuthenticated() を確認
6. 未認証なら /login に戻し、戻り先を state.from に保持
```

### 注意
- 雛形では `demo / demo` の固定認証だが、**本実装では必ず DB + bcrypt 等で照合**する
- トークンは **sessionStorage**（localStorage は使わない：XSS 経由の永続漏洩を避ける）
- パスワードはログ出力・URL クエリに絶対残さない

---

## 6. CORS と環境変数

### 6.1 CORS（`src/features/web/app.py`）
開発時は `http://localhost:5173`（Vite dev）と `http://127.0.0.1:5173` を許可。
本番ドメインを追加する場合は `allow_origins` リストを増やすか、環境変数で切替える。

### 6.2 環境変数（フロント）
| 変数名 | 既定値 | 用途 |
|---|---|---|
| `VITE_API_BASE_URL` | `/api` | API ベース URL。`.env.local` で `http://localhost:8000/api` に上書き |
| `VITE_APP_NAME` | （未設定） | 画面ヘッダーに表示するアプリ名 |
| `VITE_BASE_PATH` | `/` | サブパス配信（例：`/myapp/`）対応 |

---

## 7. ビルド・配信

### 7.1 フロントの本番ビルド
```bash
cd resources/web_template
npm run build    # dist/ に成果物が出力される
```

### 7.2 Python から静的配信したい場合
`web/app.py` で `app.mount("/", StaticFiles(directory="resources/web_template/dist", html=True))` を追加。
（テンプレートには未実装。プロジェクトの配信要件に応じて追加すること）

---

## 8. PR 前チェックリスト（Web 連携時）

- [ ] Python と Web の feature が**対で**追加されている（片側だけになっていない）
- [ ] フロントは `apiFetch` 経由で通信している（直接 `fetch` なし）
- [ ] Python の API は `service.py` を呼ぶだけの薄いラッパーになっている
- [ ] OpenAPI（`/docs`）でエンドポイントが正しく表示される
- [ ] CORS 設定に本番ドメインが含まれている（本番デプロイ時）
- [ ] `.env.local` が `.gitignore` 配下にある
- [ ] 認証トークンが sessionStorage に保存されている
- [ ] `npm run lint` / `npm run test` / `pytest` がすべて pass
