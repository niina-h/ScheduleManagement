# 共通UI・レイアウト設計ガイド

- 目的：本アプリ（上屋図面解析システム）のフロント画面構成を**標準ベース**として、今後作成する他のアプリでも同じ作法で開発・流用できるようにする。
- 対象：新規アプリのフロント実装者・レビュー担当。
- 方針：本ガイドは**現行コードの設計を文書化**したもの。コードは現状維持。新規アプリは本構成を踏襲する。
- 技術スタック：React + TypeScript + Vite + Tailwind CSS（Atomic Design）。スタイルは**Tailwind ユーティリティのみ**（インラインスタイル禁止）。

---

## 1. 全体構成（レイヤー）

```
ページ (src/pages/*)            … 画面ごとの実装。Shell に中身を流し込む
   │ 使う
レイアウト (src/components/layout/*) … 画面の外枠（Header/パンくず/本体/Footer）
   │ 使う
ドメイン部品 (src/components/domain/*) … 業務固有の複合部品（organisms）
   │ 使う
UI部品 (src/components/ui/*)     … 汎用の最小部品（atoms：Button/Card/Input…）
   │ 使う
デザイントークン (tailwind.config.js) … 色・字形・角丸・影（全レイヤーの土台）
```

- 上位は下位を使う一方向依存。**UI部品はドメイン知識を持たない**（どのアプリでも使える）。
- 認証・ルーティングは横断機能（`AuthGuard` + `router.tsx`）。

---

## 2. ディレクトリ構成と責務（標準）

```
web/src/
├─ components/
│  ├─ ui/        … 汎用UI部品（atoms）。Badge/Button/Card/Checkbox/Field/Input/
│  │              Modal/PagePlaceholder/Progress/Select/Toggle 等。barrel: index.ts
│  ├─ layout/    … レイアウトシェル。AppShell/NavigatableShell/AdminShell/
│  │              Header/Footer/Breadcrumb。barrel: index.ts
│  ├─ domain/    … 業務固有の複合部品（organisms）。★アプリごとに作り替え
│  └─ AuthGuard.tsx … 認証必須ルートの保護
├─ pages/        … 画面（ログイン以外は Shell の中に描画）
├─ lib/          … 通信・状態・ユーティリティ（api_client/auth/cn …）
├─ index.css     … グローバルベース（フォント・背景・フォーカスリング）
└─ router.tsx    … ルート定義（AuthGuard でガード）
```

各ファイル先頭には**ファイルヘッダー（ファイル名・概要）**を必ず記載（社内規約 CLAUDE.md 準拠）。

---

## 3. デザイントークン（`tailwind.config.js`）

色・字形・角丸・影は**ハードコード禁止**。`theme.extend` のトークンを Tailwind クラスで使う。

### 色（抜粋）
| 用途 | トークン（クラス例） | 値 |
| --- | --- | --- |
| ブランド | `brand` / `brandDark` / `brandDarker` / `brandTint` | #4A90D9 / #1F4E96 / #163A72 / #EAF3FB |
| 背景・面・境界 | `bg` / `surface` / `border` / `divider` | #F5F7FA / #FFFFFF / #E2E6EC / #F0F2F5 |
| テキスト階調 | `ink` / `text` / `textSub` / `textMuted` | #1F2A37 / #374151 / #6B7280 / #9CA3AF |
| ステータス | `success` / `warning` / `danger` / `info`（+`*Bg`） | 緑 / 橙 / 赤 / 青 |

使用例：`<div className="bg-surface text-ink border border-border rounded-md">`

### 字形・角丸・影
- フォント：`font-sans`（Yu Gothic 系・本文 13px は `index.css` で既定）／`font-mono`
- 角丸：`rounded-sm(3px) / rounded(4px) / rounded-md(6px) / rounded-lg(8px)`
- 影：`shadow-modal`（モーダル）／`shadow-browser`（カード/枠）

> **流用ポイント**：新規アプリは原則このトークンをそのまま使う。アプリ固有のブランド色が必要なら `brand*` の値だけ差し替える（クラス名は変えない）。

---

## 4. UI部品（atoms・`components/ui/`）

barrel から import：`import { Button, Card, Input } from "@/components/ui"`

| 部品 | 用途 |
| --- | --- |
| `Button` | ボタン（バリアント） |
| `Card` | カード枠（`shadow-browser` 系） |
| `Input` / `Select` / `Checkbox` / `Toggle` | フォーム入力 |
| `Field` | ラベル＋入力＋エラーの組（フォーム行） |
| `Badge` | ステータス表示 |
| `Modal` | モーダル（`shadow-modal`） |
| `Progress` | 進捗バー |
| `PagePlaceholder` | 空状態/準備中表示 |

ルール：
- **UI部品は業務語を含めない**（汎用名のまま）。色は props/className でトークン指定。
- 新しい汎用部品を作ったら `ui/index.ts` に export を追加。
- クラス結合は `lib/cn.ts` の `cn()` を使う（`undefined/false` を除外して結合）。

---

## 5. レイアウト（`components/layout/`）— 汎用とアプリ固有の線引き

### 5.1 AppShell（汎用・そのまま流用可）
画面の外枠を組む土台。`Header` + `Breadcrumb` + `main` + `Footer` を並べるだけ。

```tsx
<AppShell active={tab} breadcrumbs={["ホーム", "一覧"]} user={user} initials={ini} onTabClick={...}>
  {/* 画面本体 */}
</AppShell>
```
- props：`active`（アクティブタブ）/`breadcrumbs`/`user`/`initials`/`onTabClick`/`children`
- **構造自体はどのアプリでも共通**。流用可。

### 5.2 NavigatableShell（★アプリ固有・作り替え対象）
`AppShell` をラップし、**URLからアクティブタブを自動判定**＋タブ遷移＋（本アプリでは）解析ジョブの進行中バナーを表示する。

このファイルには**アプリ固有の定義**が入る：
- `TabKey`（`home/history/analyze/help/admin`）= このアプリのタブ集合
- `detectActiveTab(pathname)` = URL→タブの対応
- `tabRoute(key)` = タブ→遷移先URL
- ジョブ進行中バナー（このアプリ固有機能）

> **流用ポイント**：新規アプリでは、AppShell はそのまま使い、**NavigatableShell をアプリのタブ構成・固有バナーに合わせて作り直す**（または不要なら AppShell を直接使う）。

### 5.3 AdminShell（管理画面用の外枠）
管理系画面の共通枠。一般画面と分けたい場合に使用。

### 5.4 Header / Footer / Breadcrumb
- `Header`：ロゴ・タブ・ユーザー表示。**ロゴ/アプリ名はアプリ固有**（差し替え）。
- `Breadcrumb`：`items: string[]` を渡すだけ。汎用。
- `Footer`：汎用。

---

## 6. 画面（ページ）の作り方

1. `pages/` にコンポーネントを作る（PascalCase・ファイルヘッダー必須）。
2. ログイン以外は **Shell の中に本体を描画**する：
   ```tsx
   function Upload() {
     return (
       <NavigatableShell breadcrumbs={["ホーム", "アップロード"]}>
         {/* Card / Field / Button などの ui 部品で本体を構成 */}
       </NavigatableShell>
     )
   }
   ```
3. 複数画面で再利用する複合UIは `components/domain/`（organisms）へ切り出す。
4. アクセシビリティ（WCAG 2.1 AA）：ARIA属性・キーボード操作・フォーカス管理を実装時に確認（フォーカスリングは `index.css` で共通化済み）。

Atomic Design 階層：`ui`(atoms) → 画面内の小組（molecules）→ `domain`(organisms) → `pages`。

---

## 7. 認証・ルーティング

### AuthGuard（汎用・そのまま流用可）
```tsx
<Route path="/upload" element={<AuthGuard><Upload /></AuthGuard>} />
```
- 未認証なら `/login` へリダイレクト（戻り先を `state.from` に保持）。
- 認証状態は `lib/auth.ts` の `isAuthenticated()` で判定。

### router.tsx
- `BrowserRouter`（`App.tsx`）の中で `<Routes>` を定義。
- 保護ページは `AuthGuard` でラップ。`*` は `NotFound`。
- **サブパス配信**（例 `/uwayakaiseki/`）に対応する場合は `BrowserRouter basename` を環境変数で切替（本アプリ実装済み）。

---

## 8. 命名・コーディング規約（CLAUDE.md 準拠）

- ファイル先頭に**ファイルヘッダー**（`// ファイル名：` / `// 概要：`）。
- コンポーネント＝PascalCase、関数/変数＝camelCase、定数＝大文字スネーク。
- UI要素接頭辞（必要に応じ）：`btn` / `txt` / `lbl`、コレクション `col`。
- **スタイルは Tailwind ユーティリティのみ**。インラインスタイル・生のhex直書きは避け、トークンを使う。
- barrel（`index.ts`）で export を集約し、`@/components/ui` 形式で import。
- Server/Client の区分やコメントは日本語で明記。

---

## 9. 新規アプリ立ち上げチェックリスト（流用手順）

新しいアプリを作るときは、本アプリから**汎用層をコピー**し、**固有層を差し替え**る。

### そのままコピーして流用（汎用層）
- [ ] `tailwind.config.js`（デザイントークン）＋ `postcss.config.js`
- [ ] `src/index.css`（ベース：フォント・背景・フォーカスリング）
- [ ] `src/components/ui/`（汎用UI部品一式）＋ `index.ts`
- [ ] `src/lib/cn.ts`（クラス結合）
- [ ] `src/components/AuthGuard.tsx`＋ `src/lib/auth.ts`（認証ガード）
- [ ] `src/components/layout/{AppShell,Header,Footer,Breadcrumb}.tsx`（外枠の骨格）
- [ ] `App.tsx` / `router.tsx` の雛形（BrowserRouter＋ガード＋NotFound）

### アプリごとに作り替え（固有層）
- [ ] `NavigatableShell.tsx`：タブ集合（`TabKey`/`detectActiveTab`/`tabRoute`）と固有バナー
- [ ] `Header` のロゴ・アプリ名（必要ならブランド色 `brand*` の値）
- [ ] `components/domain/`：そのアプリの業務複合部品
- [ ] `pages/`：画面群／`router.tsx` のルート定義
- [ ] `lib/api_client.ts`：API ベース・エンドポイント

### 確認
- [ ] ファイルヘッダー・命名規約・Tailwind-only を満たす
- [ ] ログイン→主要画面の遷移が Shell 内で正しく動く
- [ ] アクセシビリティ（フォーカス・ARIA・キーボード）

---

## 10. まとめ（流用の考え方）

> **「デザイントークン＋ui/＋AppShell＋AuthGuard」＝どのアプリでも共通の土台**。
> **「NavigatableShellのタブ定義＋domain/＋pages/＋routes」＝アプリ固有**。
>
> 新規アプリは前者をコピーし、後者だけ作る。これで全アプリの見た目・操作・構造が標準化され、保守も学習コストも下がる。
