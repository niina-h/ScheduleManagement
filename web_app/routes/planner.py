"""ガントチャート入力ルート（視覚的にバーを描いてタスクを登録する方式）。

コンセプト:
- ガント風のタイムライン上で、タスクを入力し「バーをドラッグで描いて」追加する。
- 空きトラックを左→右にドラッグ＝新規タスクの期間を描画。
- 既存バーは中央ドラッグで移動、左右端ドラッグで期間伸縮。
- 保存で project_task に反映（新規は add_project_task、日付変更は update_project_task を流用）。

権限:
- 一般ユーザーは自分自身が担当者/メンバーのタスクのみ閲覧・操作できる。
- 管理職・所属長は auth_helpers.can_access_user が許可する配下ユーザーの範囲まで操作できる。
- システム管理者（マスタ）は全ユーザーを対象にできる。
- 親（レベル0）タスクの作成・編集・削除は管理職以上のみ可能。
"""
from __future__ import annotations

import logging
from datetime import date, timedelta
from typing import Any

from flask import (
    Blueprint,
    abort,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)

from ..auth_helpers import can_access_user, is_master, is_privileged
from ..models import (
    add_project_task,
    delete_project_task,
    get_all_project_tasks,
    get_all_users,
    get_user_affiliations,
    get_company_holidays,
    get_project_task_by_id,
    get_user_by_id,
    import_migration_excel,
    reassign_project_task_order,
    set_project_task_members,
    set_project_task_parent,
    update_project_task,
)

planner_bp = Blueprint("planner_bp", __name__, url_prefix="/planner")
logger = logging.getLogger(__name__)

_WEEKDAY_JA = ["月", "火", "水", "木", "金", "土", "日"]
_DISPLAY_DAYS = 63  # 画面表示日数（約9週）
_EXPORT_DAYS = 28   # Excel出力日数（4週間）
_SHIFT_DAYS = 7     # 前後ボタンの移動量（前週・次週）
_STATUSES = {"未着手", "着手", "順調", "遅れ", "完了", "停止"}


def _display_name(name: str, last_name: str | None) -> str:
    """一覧表示用の担当者名（姓があれば姓、なければ氏名）を返す。"""
    return (last_name or name or "").strip()


def _parse_ids(s: Any) -> list[int]:
    """カンマ区切りID文字列を int リストに変換する。"""
    out: list[int] = []
    for part in str(s or "").split(","):
        part = part.strip()
        if part.isdigit():
            out.append(int(part))
    return out


def _members(t: dict) -> list[int]:
    """タスク（親）の関係ユーザー（メンバー）ID一覧を返す。"""
    return _parse_ids(t.get("member_ids"))


def _get_visible_tree(
    login_user_id: int, login_role: str, target_user_id: int
) -> tuple[list[dict], dict[int, list[dict]]]:
    """役職に応じて閲覧可能な親タスク（ルート）と子マップを返す。

    マスタ=全件（対象ユーザー指定時はその関係者のみ）、管理職=自分がメンバーの親のみ、
    一般=自分が担当の子を含む親ツリーのみを返す。

    Args:
        login_user_id: ログインユーザーID。
        login_role: ログインユーザーの役職。
        target_user_id: 絞り込み対象ユーザーID（0=全員）。

    Returns:
        tuple[list[dict], dict[int, list[dict]]]: (可視ルート一覧, 親ID→子タスク一覧)
    """
    master = is_master(login_role)
    manager = login_role == "管理職"

    # 全タスク（イベント・マイルストーン・完了/停止も含む）をツリー化して表示する。
    # 期間（開始日・終了日）が無いものはバーを描けないため除外する。
    raw = [
        t for t in get_all_project_tasks()
        if (t.get("start_date") or "").strip() and (t.get("end_date") or "").strip()
    ]
    by_id = {t["id"]: t for t in raw}
    children: dict[int, list[dict]] = {}
    roots: list[dict] = []
    for t in raw:
        pid = t.get("parent_task_id")
        if pid and pid in by_id:
            children.setdefault(pid, []).append(t)
        else:
            roots.append(t)

    def _subtree_involves(root: dict, uid: int) -> bool:
        """ルート配下（自身含む）に uid がメンバー/担当として関与するか。"""
        if uid in _members(root):
            return True
        stack = [root]
        while stack:
            n = stack.pop()
            if uid in (n.get("assigned_to"), n.get("assigned_to_2")):
                return True
            stack.extend(children.get(n["id"], []))
        return False

    if master:
        if target_user_id == 0:
            visible_roots = list(roots)
        else:
            visible_roots = [r for r in roots if _subtree_involves(r, target_user_id)]
    elif manager:
        # 管理職：自分が関係ユーザー（メンバー）の親のみ
        visible_roots = [r for r in roots if login_user_id in _members(r)]
    else:
        # 一般：自分が担当の子を含む親ツリー
        visible_roots = [r for r in roots if _subtree_involves(r, login_user_id)]

    return visible_roots, children


def _monday(d: date) -> date:
    """指定日を含む週の月曜日を返す。"""
    return d - timedelta(days=d.weekday())


def _default_range_start() -> date:
    """デフォルトの表示開始日（先週の月曜）。"""
    return _monday(date.today()) - timedelta(days=7)


@planner_bp.before_request
def _require_login() -> Any:
    """未ログインならログイン画面へ誘導する。"""
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))
    return None


def _resolve_target(login_user_id: int) -> int:
    """?user_id= から対象ユーザーIDを決定する。

    - 戻り値 0 は「全員」を表す（マスタのみ許可）。
    - 既定値はマスタなら全員(0)、それ以外は本人。
    - 権限外のユーザーが指定された場合は既定値へフォールバックする
      （一般は自分自身のみ、管理職・所属長は can_access_user が許可する範囲のみ）。
    """
    login_role = session.get("user_role", "")
    default = 0 if is_master(login_role) else login_user_id
    req_uid = request.args.get("user_id", "").strip()
    if not req_uid:
        return default
    try:
        cand = int(req_uid)
    except ValueError:
        return default
    if cand == 0:
        return 0 if is_master(login_role) else default
    if cand == login_user_id:
        return cand
    if not is_privileged(login_role):
        return default
    target = get_user_by_id(cand)
    if not target:
        return default
    login_user_dict = {
        "id": login_user_id, "role": login_role, "dept": session.get("user_dept", ""),
    }
    return cand if can_access_user(login_user_dict, target) else default


@planner_bp.route("/")
def planner() -> Any:
    """ガント入力（試作）画面を表示する。"""
    login_user_id = int(session["user_id"])
    target_user_id = _resolve_target(login_user_id)

    # 表示開始日
    raw_start = request.args.get("start", "").strip()
    try:
        range_start = date.fromisoformat(raw_start) if raw_start else _default_range_start()
    except ValueError:
        range_start = _default_range_start()
    range_end = range_start + timedelta(days=_DISPLAY_DAYS - 1)
    start_iso = range_start.isoformat()
    end_iso = range_end.isoformat()
    today_iso = date.today().isoformat()

    # 会社休日（祝日含む）の日付集合。土日祝の背景色判定に使う。
    holiday_dates: set[str] = {
        h["holiday_date"] for h in get_company_holidays()
    }

    # 日付ヘッダー
    days: list[dict] = []
    for i in range(_DISPLAY_DAYS):
        d = range_start + timedelta(days=i)
        iso = d.isoformat()
        days.append({
            "index": i,
            "iso": iso,
            "m": d.month,
            "d": d.day,
            "wd": _WEEKDAY_JA[d.weekday()],
            "weekend": d.weekday() >= 5,
            "holiday": iso in holiday_dates,
            "is_today": iso == today_iso,
            "is_month_start": d.day == 1,
        })

    # ── 役職 ──
    login_role = session.get("user_role", "")
    master = is_master(login_role)
    manager = login_role == "管理職"
    privileged = master or manager  # 親タスクの作成・編集が可能

    # 全ユーザー（ID→表示名）。既存タスクのメンバー・担当名の表示に使う（部署に関わらず引ける必要がある）。
    users = get_all_users()
    user_name_map = {
        u["id"]: _display_name(u.get("name", ""), u.get("last_name")) for u in users
    }
    # 担当・メンバーの選択候補は実効所属に限定する
    # （システム管理者が所属切替中はその所属、通常は自分の所属）。
    login_dept = session.get("user_dept", "")
    # 記憶トークン自動ログインの古いセッションは user_dept が空のことがあるためDBから補完。
    if not login_dept:
        _lu = get_user_by_id(login_user_id)
        if _lu:
            login_dept = _lu.get("dept", "") or ""
            session["user_dept"] = login_dept
    from ..auth_helpers import (
        get_active_scope_dept, get_effective_depts, is_dept_head, is_system_admin,
    )
    scope_dept = get_active_scope_dept(login_role, login_dept, session.get("active_dept"))
    # 候補ユーザーの所属集合を役職別に決める（所属長は担当所属のみ・他所属を出さない）。
    if is_dept_head(login_role):
        _depts = get_effective_depts(
            {"role": login_role, "dept": login_dept},
            affiliations=get_user_affiliations(login_user_id),
            active_dept=session.get("active_dept"),
        ) or []
        _allow = set(_depts)
        assign_users = [
            {"id": u["id"], "name": user_name_map[u["id"]]}
            for u in users if u.get("dept") in _allow
        ]
    elif is_system_admin(login_role) and not scope_dept:
        assign_users = [
            {"id": u["id"], "name": user_name_map[u["id"]]} for u in users
        ]
    else:
        assign_users = [
            {"id": u["id"], "name": user_name_map[u["id"]]}
            for u in users if u.get("dept") == scope_dept
        ]

    visible_roots, children = _get_visible_tree(login_user_id, login_role, target_user_id)

    def _sortkey(t: dict) -> tuple:
        # 手動並べ替え（display_order）を最優先。同順位は開始日→IDで安定化。
        return (t.get("display_order") or 0, (t.get("start_date") or "9999-99-99"), t["id"])

    def _people_list(ids: list[int]) -> list[dict]:
        return [{"id": i, "name": user_name_map.get(i, "?")} for i in ids]

    tasks: list[dict] = []

    def _assignees(t: dict) -> list[dict]:
        result: list[dict] = []
        if t.get("assigned_to"):
            result.append({"id": t["assigned_to"],
                           "name": _display_name(t.get("assigned_name", ""), t.get("assigned_last_name"))})
        if t.get("assigned_to_2"):
            result.append({"id": t["assigned_to_2"],
                           "name": _display_name(t.get("assigned_name_2", ""), t.get("assigned_last_name_2"))})
        return result

    def _emit(t: dict, level: int) -> None:
        # 「親」は level ではなく「実際に子を持つか」で判定する（保存側と統一）。
        is_parent = bool(children.get(t["id"]))
        # 一般ユーザーは最上位の親タスクを編集不可（子・単独タスクは可）。
        editable = privileged or not (is_parent and level == 0)
        tasks.append({
            "id": t["id"],
            "task_name": t.get("task_name", ""),
            "start": (t.get("start_date") or "").strip(),
            "end": (t.get("end_date") or "").strip(),
            "status": t.get("status") or "",
            "progress": t.get("progress") or 0,
            "delay_days": t.get("delay_days") or 0,
            # 親（子持ち）はメンバー、それ以外（単独・末端）は担当者を people として渡す。
            "people": _people_list(_members(t)) if is_parent else _assignees(t),
            "level": level,
            "parent_id": t.get("parent_task_id") or None,
            "editable": editable,
        })
        for c in sorted(children.get(t["id"], []), key=_sortkey):
            _emit(c, level + 1)

    for r in sorted(visible_roots, key=_sortkey):
        _emit(r, 0)

    # 人物選択（上部プルダウン）はマスタのみ表示：全員 ＋ 全ユーザー。
    people = [{"id": 0, "name": "全員"}] + assign_users

    lu = get_user_by_id(login_user_id) or {}
    login_user = {"id": login_user_id,
                  "name": _display_name(lu.get("name", ""), lu.get("last_name"))}
    # 新規タスクの既定担当者：マスタが特定ユーザー閲覧中はその人、それ以外は操作者本人。
    if master and target_user_id not in (0, login_user_id):
        u = get_user_by_id(target_user_id) or {}
        default_assignee = {"id": target_user_id,
                            "name": _display_name(u.get("name", ""), u.get("last_name"))}
        view_name = default_assignee["name"]
    else:
        default_assignee = login_user
        view_name = "全員" if (master and target_user_id == 0) else login_user["name"]

    return render_template(
        "gantt_input_test.html",
        days=days,
        range_start=start_iso,
        range_end=end_iso,
        today_iso=today_iso,
        prev_start=(range_start - timedelta(days=_SHIFT_DAYS)).isoformat(),
        next_start=(range_start + timedelta(days=_SHIFT_DAYS)).isoformat(),
        this_start=_default_range_start().isoformat(),
        tasks=tasks,
        people=people,
        selected_user_id=target_user_id,
        assign_users=assign_users,
        default_assignee=default_assignee,
        login_user=login_user,
        view_name=view_name,
        show_selector=master,
        can_parent=privileged,
        role=login_role,
        is_master=master,
        display_days=_DISPLAY_DAYS,
        csrf_token=session.get("csrf_token", ""),
    )


@planner_bp.route("/export")
def export() -> Any:
    """ガントチャートをExcel出力する（親=大項目、子=タスクの階層で記載）。

    画面で見えている範囲（役職別の可視ツリー・表示期間）と一致させるため、
    タスク管理のガントExcel（大区分/中区分グルーピング）とは別に、
    親子ツリー専用の出力関数を使う。
    """
    from flask import send_file
    from .project_tasks import _build_gantt_excel_by_tree

    login_user_id = int(session["user_id"])
    login_role = session.get("user_role", "")
    target_user_id = _resolve_target(login_user_id)

    raw_start = request.args.get("start", "").strip()
    try:
        range_start = date.fromisoformat(raw_start) if raw_start else _default_range_start()
    except ValueError:
        range_start = _default_range_start()

    visible_roots, children = _get_visible_tree(login_user_id, login_role, target_user_id)
    # Excel出力は4週間（28日）分のみ出力する（画面表示の9週とは別）。
    wb = _build_gantt_excel_by_tree(visible_roots, children, range_start, _EXPORT_DAYS)

    import io
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    today_str = date.today().isoformat()
    return send_file(
        buf, as_attachment=True, download_name=f"ガントチャート_{today_str}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@planner_bp.route("/export-pdf")
def export_pdf() -> Any:
    """ガントチャートをPDF出力する（Excel出力と同じ内容・レイアウトをPDF化）。

    Excel出力（/export）と同じ可視ツリー・同じ4週間を使い、reportlabで描画する。
    """
    from flask import send_file
    from .project_tasks import _build_gantt_pdf_by_tree

    login_user_id = int(session["user_id"])
    login_role = session.get("user_role", "")
    target_user_id = _resolve_target(login_user_id)

    raw_start = request.args.get("start", "").strip()
    try:
        range_start = date.fromisoformat(raw_start) if raw_start else _default_range_start()
    except ValueError:
        range_start = _default_range_start()

    visible_roots, children = _get_visible_tree(login_user_id, login_role, target_user_id)
    # PDFもExcelと同じ4週間（28日）分を出力する。
    buf = _build_gantt_pdf_by_tree(visible_roots, children, range_start, _EXPORT_DAYS)

    today_str = date.today().isoformat()
    return send_file(
        buf, as_attachment=True, download_name=f"ガントチャート_{today_str}.pdf",
        mimetype="application/pdf",
    )


@planner_bp.route("/save", methods=["POST"])
def save() -> Any:
    """ガント入力の内容を project_task に保存する。

    JSON: {
      user_id,
      new_tasks: [{task_name, start, end, category_id?, subcategory_id?}],
      updated:   [{id, start, end}]
    }
    """
    if request.headers.get("X-CSRF-Token", "") != session.get("csrf_token"):
        return jsonify({"ok": False, "error": "CSRFトークン不正"}), 400

    login_user_id = int(session["user_id"])
    login_role = session.get("user_role", "")
    payload = request.get_json(silent=True) or {}

    # 対象ユーザー。0（全員）はマスタのみ許可。他人を指定できるのは
    # can_access_user が許可する範囲（管理職・所属長は配下、マスタは全員）のみで、
    # 一般ユーザーは自分自身以外を指定できない。
    target_user_id = login_user_id
    req_uid = payload.get("user_id")
    if req_uid not in (None, ""):
        try:
            cand = int(req_uid)
        except (TypeError, ValueError):
            return jsonify({"ok": False, "error": "ユーザー指定が不正です"}), 400
        if cand == 0:
            if not is_master(login_role):
                return jsonify({"ok": False, "error": "権限がありません"}), 403
            target_user_id = 0
        elif cand == login_user_id:
            target_user_id = cand
        else:
            target = get_user_by_id(cand)
            if not target:
                return jsonify({"ok": False, "error": "ユーザーが見つかりません"}), 400
            login_user_dict = {
                "id": login_user_id, "role": login_role, "dept": session.get("user_dept", ""),
            }
            if not is_privileged(login_role) or not can_access_user(login_user_dict, target):
                return jsonify({"ok": False, "error": "権限がありません"}), 403
            target_user_id = cand
    default_owner = target_user_id if target_user_id else login_user_id

    def _valid_date(v: Any) -> str | None:
        try:
            return date.fromisoformat(str(v)).isoformat()
        except (TypeError, ValueError):
            return None

    def _parse_status(v: Any) -> str:
        """状態文字列を検証（不正値は未着手）。"""
        return v if v in _STATUSES else "未着手"

    def _parse_int(v: Any, lo: int = 0, hi: int | None = None) -> int:
        """整数化してlo/hiにクランプする（不正値は0扱い）。"""
        try:
            n = int(v)
        except (TypeError, ValueError):
            n = 0
        n = max(lo, n)
        return min(hi, n) if hi is not None else n

    def _parse_people(row: dict, limit: int | None = None) -> list[int]:
        """行の関係ユーザー/担当者ID（実在ユーザーのみ、任意で人数制限）を返す。"""
        result: list[int] = []
        for a in (row.get("people") or []):
            try:
                av = int(a)
            except (TypeError, ValueError):
                continue
            if av not in result and get_user_by_id(av):
                result.append(av)
            if limit is not None and len(result) >= limit:
                break
        return result

    # 役職による権限：親（レベル0）の作成・編集は管理職/マスタのみ。
    privileged = is_master(login_role) or login_role == "管理職"

    login_user_dict = {
        "id": login_user_id, "role": login_role, "dept": session.get("user_dept", ""),
    }

    def _can_touch_existing(existing: dict) -> bool:
        """既存タスクの更新・削除をログインユーザーが行えるか判定する。

        マスタは常に可。それ以外は、タスクの担当者・メンバーに自分が含まれるか、
        またはそのいずれかに can_access_user でアクセス可能な場合のみ許可する。
        """
        if is_master(login_role):
            return True
        related_ids = set(_parse_ids(existing.get("member_ids")))
        for key in ("assigned_to", "assigned_to_2"):
            v = existing.get(key)
            if v:
                related_ids.add(v)
        if login_user_id in related_ids:
            return True
        if not is_privileged(login_role):
            return False
        for uid in related_ids:
            u = get_user_by_id(uid)
            if u and can_access_user(login_user_dict, u):
                return True
        return False

    created = 0
    updated = 0
    tmp_map: dict[str, int] = {}   # クライアント一時ID → 実タスクID
    ordered_ids: list[int] = []    # DOM順（保存対象の実タスクID）— 表示順の再割当に使う

    # rows は親→子の順（DOM順）で受け取る。親IDは既存ID or 一時ID で参照。
    src_rows: list[dict] = payload.get("rows", []) or []

    # 「親」判定は level==0 ではなく「実際に子を持つか」で行う。
    # DOM順のため、次以降で自分よりレベルが大きい行が現れれば子を持つ親とみなす
    # （同レベル以下が先に現れたら子はいない）。子を持たない最上位行は担当タスクとして扱い、
    # 担当者を設定して週間予定へ反映できるようにする。
    def _has_children(idx: int) -> bool:
        lv = _parse_int(src_rows[idx].get("level"), 0)
        for j in range(idx + 1, len(src_rows)):
            nxt_lv = _parse_int(src_rows[j].get("level"), 0)
            if nxt_lv > lv:
                return True
            if nxt_lv <= lv:
                return False
        return False

    for row_idx, row in enumerate(src_rows):
        name = str(row.get("name", "") or "").strip()
        s = _valid_date(row.get("start"))
        e = _valid_date(row.get("end"))
        if s and e and e < s:
            s, e = e, s
        level = _parse_int(row.get("level"), 0)
        # 子を持つ行のみ「親（メンバーのみ・担当なし）」。それ以外は担当タスク。
        is_parent = _has_children(row_idx)

        # 一般ユーザーは最上位（レベル0）の親タスクを作成・編集できない。
        if is_parent and level == 0 and not privileged:
            continue

        # 親IDの解決（既存 parent_id 優先、なければ一時ID）
        parent_real: int | None = None
        if row.get("parent_id") not in (None, "", "null"):
            try:
                parent_real = int(row["parent_id"])
            except (TypeError, ValueError):
                parent_real = None
        elif row.get("parent_tmp") and row["parent_tmp"] in tmp_map:
            parent_real = tmp_map[row["parent_tmp"]]

        rid = row.get("id")
        if rid in (None, "", "null"):
            # ── 新規タスク（バー未描画＝日付なしはスキップ） ──
            if not name or not s or not e:
                continue
            if is_parent:
                # 親：メンバー（関係ユーザー）を保存。担当者は持たない。
                member_ids = _parse_people(row)
                a1, a2, mem_str = None, None, ",".join(str(i) for i in member_ids)
            else:
                # 子：担当者（最大2名）。未指定は既定担当者。
                a_ids = _parse_people(row, 2) or [default_owner]
                a1 = a_ids[0]
                a2 = a_ids[1] if len(a_ids) > 1 else None
                mem_str = ""
            new_id = add_project_task(
                category_id=None, subcategory_id=None,
                task_name=name, description="",
                start_date=s, end_date=e,
                status=_parse_status(row.get("status")),
                progress=_parse_int(row.get("progress"), 0, 100),
                delay_days=_parse_int(row.get("delay"), 0),
                created_by=login_user_id,
                updated_by=session.get("user_name", ""),
                assigned_to=a1, assigned_to_2=a2,
                member_ids=mem_str,
            )
            if parent_real:
                set_project_task_parent(new_id, parent_real)
            if row.get("tmp"):
                tmp_map[str(row["tmp"])] = new_id
            created += 1
            ordered_ids.append(new_id)
        else:
            # ── 既存タスク ──
            try:
                tid = int(rid)
            except (TypeError, ValueError):
                continue
            existing = get_project_task_by_id(tid)
            if not existing:
                continue
            # 子タスクの親参照解決用に、非変更行でも一時IDを登録し、表示順に含める。
            tmp_map[str(row.get("tmp"))] = tid
            ordered_ids.append(tid)
            if not row.get("dirty"):
                continue
            if not _can_touch_existing(existing):
                continue
            if is_parent:
                # 親：メンバーを更新。担当者は変更しない。
                set_project_task_members(tid, ",".join(str(i) for i in _parse_people(row)))
                a1 = existing.get("assigned_to")
                a2 = existing.get("assigned_to_2")
            else:
                a_ids = _parse_people(row, 2)
                a1 = a_ids[0] if a_ids else existing.get("assigned_to")
                a2 = (a_ids[1] if len(a_ids) > 1 else None) if a_ids else existing.get("assigned_to_2")
            if s and e:
                update_project_task(
                    task_id=tid,
                    category_id=existing.get("category_id"),
                    subcategory_id=existing.get("subcategory_id"),
                    task_name=existing.get("task_name", ""),
                    description=existing.get("description", "") or "",
                    start_date=s, end_date=e,
                    status=_parse_status(row.get("status")),
                    progress=_parse_int(row.get("progress"), 0, 100),
                    delay_days=_parse_int(row.get("delay"), 0),
                    updated_by=session.get("user_name", ""),
                    assigned_to=a1,
                    assigned_to_2=a2,
                    is_milestone=existing.get("is_milestone", 0) or 0,
                    is_event=existing.get("is_event", 0) or 0,
                    planned_hours=existing.get("planned_hours", 0.0) or 0.0,
                )
            set_project_task_parent(tid, parent_real)
            updated += 1

    # ── 削除（クライアントで削除された既存タスク） ──
    # 一般ユーザーは親（最上位）タスクを削除できない。
    deleted = 0
    for d in payload.get("deleted", []) or []:
        try:
            did = int(d)
        except (TypeError, ValueError):
            continue
        existing = get_project_task_by_id(did)
        if not existing:
            continue
        if not existing.get("parent_task_id") and not privileged:
            continue  # 親タスクは一般ユーザー削除不可
        if not _can_touch_existing(existing):
            continue
        delete_project_task(did)
        deleted += 1

    # DOM順に表示順（display_order）を再割当し、上下移動・挿入を永続化する。
    reassign_project_task_order(ordered_ids)

    return jsonify({"ok": True, "created": created, "updated": updated, "deleted": deleted})


@planner_bp.route("/import-excel", methods=["POST"])
def import_excel() -> Any:
    """移行用フラットExcelをアップロードし project_task へ取り込む（システム管理者のみ）。

    タスク管理一覧画面の「タスク移行Excel出力」で出力したファイルを想定する。
    id列が既存タスクと一致すれば更新、空欄なら新規追加し、parent_task_id列で
    親子関係（新ガントのツリー構造）を設定する。

    Returns:
        Any: ガントチャート画面へのリダイレクト。
    """
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    if not is_master(session.get("user_role", "")):
        flash("この操作にはシステム管理者権限が必要です", "danger")
        return redirect(url_for("planner_bp.planner"))

    uploaded = request.files.get("migration_file")
    if not uploaded or not uploaded.filename:
        flash("ファイルを選択してください", "warning")
        return redirect(url_for("planner_bp.planner"))
    if not uploaded.filename.endswith(".xlsx"):
        flash(".xlsx形式のファイルを選択してください", "warning")
        return redirect(url_for("planner_bp.planner"))

    import tempfile
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    uploaded.save(tmp.name)

    result = import_migration_excel(
        file_path=tmp.name,
        login_user_id=int(session["user_id"]),
        updated_by=session.get("user_name", ""),
    )

    parts = []
    if result["imported"]:
        parts.append(f"{result['imported']}件追加")
    if result["updated"]:
        parts.append(f"{result['updated']}件更新")
    msg = f"インポート完了: {' / '.join(parts) if parts else '変更なし'}"
    if result["errors"]:
        msg += f" / {len(result['errors'])}件エラー"
        for e in result["errors"]:
            logger.warning("移行Excelインポートエラー: %s", e)
    flash(msg, "success" if (result["imported"] > 0 or result["updated"] > 0) else "info")
    return redirect(url_for("planner_bp.planner"))
