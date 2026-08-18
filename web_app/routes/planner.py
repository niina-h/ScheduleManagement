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
    clear_project_task_assigned_to,
    delete_project_task,
    get_accessible_users,
    get_all_project_tasks,
    get_all_users,
    get_user_affiliations,
    get_company_holidays,
    get_project_task_by_id,
    get_user_by_id,
    import_migration_excel,
    is_dev_summary_enabled,
    reassign_project_task_order,
    resolve_parent_progress,
    resolve_parent_status,
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
_STATUSES = {"未着手", "着手", "順調", "遅れ", "完了", "中断"}


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
    login_user_id: int, login_role: str, target_user_id: int, login_dept: str = "",
) -> tuple[list[dict], dict[int, list[dict]]]:
    """役職に応じて閲覧可能な親タスク（ルート）と子マップを返す。

    マスタ=全件（対象ユーザー指定時はその関係者のみ）、管理職・所属長=スコープ内
    メンバー（get_accessible_users）の誰かがメンバーの親のみ、一般=自分が担当の
    子を含む親ツリーのみを返す。

    Args:
        login_user_id: ログインユーザーID。
        login_role: ログインユーザーの役職。
        target_user_id: 絞り込み対象ユーザーID（0=全員）。
        login_dept: ログインユーザーの所属（管理職・所属長のスコープ判定に使用）。

    Returns:
        tuple[list[dict], dict[int, list[dict]]]: (可視ルート一覧, 親ID→子タスク一覧)
    """
    master = is_master(login_role)
    manager = is_privileged(login_role) and not master

    # 全タスク（マイルストーン・完了/中断も含む）をツリー化して表示する。
    # 期間（開始日・終了日）が無いものはバーを描けないが、タスク名・担当者の行としては
    # 表示する（ガントバーを削除しても行自体は残せるようにするため）。
    # ただし通常のイベント（会議・打合せ等、is_event=1 かつ is_milestone=0）は
    # ガントチャートには反映しない（イベント専用画面でのみ管理する）。
    # マイルストーン（is_milestone=1）は is_event の値に関わらず表示を維持する。
    raw = [
        t for t in get_all_project_tasks()
        if not (t.get("is_event") and not t.get("is_milestone"))
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
        # 管理職・所属長：スコープ内メンバー（自分自身を含む）の誰かが、
        # 親のメンバー、または配下の子タスクの担当者として関与する親のみ表示する
        # （親の member_ids だけで判定すると、子の担当者しか設定されていない
        #   ツリーが表示から漏れてしまうため、_subtree_involves で判定する）。
        # 担当者プルダウンで個別ユーザーが指定されていれば、そのユーザーが
        # 関与する親のみにさらに絞り込む。
        scope_ids = {u["id"] for u in get_accessible_users(login_user_id, login_role, login_dept)}
        scope_ids.add(login_user_id)
        visible_roots = [r for r in roots if any(_subtree_involves(r, uid) for uid in scope_ids)]
        if target_user_id and target_user_id in scope_ids:
            visible_roots = [r for r in visible_roots if _subtree_involves(r, target_user_id)]
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


def _export_file_stem(login_dept: str) -> str:
    """ガントチャート出力ファイル名の共通部分（拡張子なし）を返す。

    形式は「所属＋進捗＋yymmdd」（ログイン中の本人の所属を使用）。
    所属が未設定の場合は「進捗」のみとする。

    Args:
        login_dept: ログインユーザーの所属名。

    Returns:
        str: ファイル名の共通部分（例: "○○部進捗260728"）。
    """
    yymmdd = date.today().strftime("%y%m%d")
    return f"{login_dept}進捗{yymmdd}" if login_dept else f"進捗{yymmdd}"


@planner_bp.before_request
def _require_login() -> Any:
    """未ログインならログイン画面へ誘導する。"""
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))
    return None


def _resolve_target(login_user_id: int) -> int:
    """?user_id= から対象ユーザーIDを決定する。

    - 戻り値 0 は「全員（スコープ内混在表示）」を表す。
    - 既定値は、権限者（管理職・所属長・マスタ）なら全員(0)、一般は本人。
    - 権限外のユーザーが指定された場合は既定値へフォールバックする
      （一般は自分自身のみ、管理職・所属長は can_access_user が許可する範囲のみ）。
    """
    login_role = session.get("user_role", "")
    privileged = is_privileged(login_role)
    default = 0 if privileged else login_user_id
    req_uid = request.args.get("user_id", "").strip()
    if not req_uid:
        return default
    try:
        cand = int(req_uid)
    except ValueError:
        return default
    if cand == 0:
        return 0 if privileged else default
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

    # 完了タスク（状態が「完了」）の表示切替。既定は非表示。
    show_done = request.args.get("show_done", "").strip() == "1"

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
    privileged = is_privileged(login_role)  # 親タスクの作成・編集が可能（管理職以上）

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
    # 日報メール＜開発状況＞の「開発集計に入れる」チェックを画面に出すかどうか
    # （部門ごとに機能自体を有効化できる。既定は無効）。
    dev_summary_enabled = is_dev_summary_enabled(scope_dept or login_dept)
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
        # 部署未設定のテスト用アカウント等が担当者候補に混入するのを避けるため、
        # 所属が設定されているユーザーのみを対象にする。
        assign_users = [
            {"id": u["id"], "name": user_name_map[u["id"]]}
            for u in users if u.get("dept")
        ]
    else:
        assign_users = [
            {"id": u["id"], "name": user_name_map[u["id"]]}
            for u in users if u.get("dept") == scope_dept
        ]

    visible_roots, children = _get_visible_tree(login_user_id, login_role, target_user_id, login_dept)

    def _sortkey(t: dict) -> tuple:
        # 手動並べ替え（display_order）を最優先。同順位は開始日→IDで安定化。
        return (t.get("display_order") or 0, (t.get("start_date") or "9999-99-99"), t["id"])

    def _people_list(ids: list[int]) -> list[dict]:
        return [{"id": i, "name": user_name_map.get(i, "?")} for i in ids]

    tasks: list[dict] = []

    def _assignees(t: dict) -> list[dict]:
        # assigned_ids（2名以上対応）が設定されていればそちらを優先する。
        # 未設定（従来データ）の場合は assigned_to/assigned_to_2 にフォールバックする。
        ids_str = (t.get("assigned_ids") or "").strip()
        if ids_str:
            return _people_list(_parse_ids(ids_str))
        result: list[dict] = []
        if t.get("assigned_to"):
            result.append({"id": t["assigned_to"],
                           "name": _display_name(t.get("assigned_name", ""), t.get("assigned_last_name"))})
        if t.get("assigned_to_2"):
            result.append({"id": t["assigned_to_2"],
                           "name": _display_name(t.get("assigned_name_2", ""), t.get("assigned_last_name_2"))})
        return result

    def _assignee_ids(t: dict) -> list[int]:
        """タスク（子・単独）の担当者ID一覧を返す（assigned_ids優先、無ければ assigned_to系）。"""
        ids_str = (t.get("assigned_ids") or "").strip()
        if ids_str:
            return _parse_ids(ids_str)
        return [i for i in (t.get("assigned_to"), t.get("assigned_to_2")) if i]

    def _flatten_descendants(t: dict) -> list[dict]:
        """このノード（自身含む）配下の全ノードをフラットなリストで返す。"""
        out: list[dict] = []
        stack = [t]
        while stack:
            n = stack.pop()
            out.append(n)
            stack.extend(children.get(n["id"]) or [])
        return out

    def _leaf_visible(t: dict, member_matched: bool = False) -> bool:
        """末端（子を持たないノード）が表示対象かどうか（完了・担当者絞り込みを判定）。

        Args:
            member_matched: 祖先のいずれかの親の member_ids に対象ユーザーが
                含まれる場合 True。この場合、子タスク自身の担当者に含まれて
                いなくても表示対象とする（親のメンバーに割り当てられていれば
                配下をすべて見られるようにするため）。
        """
        status = t.get("status") or ""
        if not show_done and status == "完了":
            return False
        if target_user_id and not member_matched and target_user_id not in _assignee_ids(t):
            return False
        return True

    _has_visible_cache: dict[tuple[int, bool], bool] = {}

    def _has_visible_descendant(t: dict, member_matched: bool = False) -> bool:
        """このノード（自身含む）の配下に、最終的に表示される末端が1つでもあるか。

        中間階層の見出し（子は持つが、配下が全員フィルタで除外された）が
        中身の無いまま表示され続けるのを防ぐため、事前にボトムアップで判定する。

        Args:
            member_matched: 祖先のいずれかの親の member_ids に対象ユーザーが
                含まれる場合 True（_leaf_visible と同様の緩和を配下に伝える）。
        """
        cache_key = (t["id"], member_matched)
        if cache_key in _has_visible_cache:
            return _has_visible_cache[cache_key]
        matched = member_matched or (bool(target_user_id) and target_user_id in _members(t))
        kids = children.get(t["id"]) or []
        if not kids:
            result = _leaf_visible(t, matched)
        else:
            # 完了ツリーの非表示判定は、親自身にバー（start_date・end_date）が
            # 設定されている場合のみ行う。バー無しの見出し行（子の状況から状態を
            # 自動判定する行）は、たまたま今の子が全員完了でも、まだ新しい子を
            # 追加できる「箱」として残したいため、子が全員完了でも消さない。
            # この場合、配下の子がすべて個別に非表示（完了）でも、対象ユーザーが
            # この親のメンバー・担当のいずれかに関与していれば親自身は表示する
            # （_emit 側の完了非表示の緩和と揺れないよう、ここでも「箱」を残す）。
            has_bar = bool((t.get("start_date") or "").strip() and (t.get("end_date") or "").strip())
            status = resolve_parent_status(t, kids)
            if has_bar and not show_done and status == "完了":
                result = False
            else:
                result = any(_has_visible_descendant(c, matched) for c in kids)
                if not result and not has_bar and target_user_id:
                    # 配下がすべて完了で個別に非表示でも、対象ユーザーがこの親の
                    # メンバー（matched）か、配下いずれかの担当者であれば、この親
                    # （箱）自体は表示する（完了タスクを全部やり切った直後、
                    # 担当者本人の画面からもツリーが消えてしまう不具合を防ぐ）。
                    result = matched or any(
                        target_user_id in _assignee_ids(c)
                        for c in _flatten_descendants(t)
                        if not (children.get(c["id"]) or [])
                    )
        _has_visible_cache[cache_key] = result
        return result

    def _emit(t: dict, level: int, member_matched: bool = False) -> None:
        # 「親」は level ではなく「実際に子を持つか」で判定する（保存側と統一）。
        kids = children.get(t["id"]) or []
        is_parent = bool(kids)
        # 一般ユーザーは最上位の親タスクを編集不可（子・単独タスクは可）。
        editable = privileged or not (is_parent and level == 0)
        # 親自身にバー（開始日・終了日）が無い見出し行は、子の状況から状態・進捗を自動判定する
        # （子に進捗があっても親が常に「未着手」表示になってしまう不具合を防ぐ）。
        status = resolve_parent_status(t, kids) if is_parent else (t.get("status") or "")
        progress = resolve_parent_progress(t, kids) if is_parent else (t.get("progress") or 0)
        # 完了タスクの表示切替：終了日に関わらず、状態が「完了」ならその行（配下の
        # 子孫も含む）を非表示にする。子の有無に関係なく判定するため、一部の子だけ
        # 間引かれてDOM順・level構造が崩れることはない（サブツリーごと消える）。
        # ただし、子を持つ親自身にバー（start_date・end_date）が無い場合（子の
        # 状況から状態を自動判定する見出し行）は、たまたま今の子が全員完了でも
        # 非表示にしない。まだ新しい子タスクを追加できる「箱」であり、子が1件だけ
        # 完了した途端にツリーごと消えると分かりにくいため。
        has_bar = bool((t.get("start_date") or "").strip() and (t.get("end_date") or "").strip())
        if (not is_parent or has_bar) and not show_done and status == "完了":
            return
        # このノード自身が親で、member_ids に対象ユーザーを含む場合、以降（自身の
        # 配下）は「メンバー一致」扱いとし、子タスクの担当者に入っていなくても
        # 表示対象にする（親のメンバーに割り当てられていれば配下をすべて見られる
        # ようにするため）。
        member_matched = member_matched or (
            bool(target_user_id) and is_parent and target_user_id in _members(t)
        )
        # 担当者プルダウンで個別ユーザーに絞り込んでいる場合：
        # - 末端（子・単独）は、そのユーザーが担当者でなく、祖先のメンバー一致も
        #   無ければ間引く。
        # - 中間階層（子を持つ）は、配下に表示対象の末端が1つも無ければ、
        #   中身の無い見出しとして残さず間引く。
        if target_user_id:
            if not is_parent and not member_matched and target_user_id not in _assignee_ids(t):
                return
            if is_parent and not _has_visible_descendant(t, member_matched):
                return
        tasks.append({
            "id": t["id"],
            "task_name": t.get("task_name", ""),
            "start": (t.get("start_date") or "").strip(),
            "end": (t.get("end_date") or "").strip(),
            "status": status,
            "progress": progress,
            "delay_days": t.get("delay_days") or 0,
            # 親（子持ち）はメンバー、それ以外（単独・末端）は担当者を people として渡す。
            "people": _people_list(_members(t)) if is_parent else _assignees(t),
            "level": level,
            "parent_id": t.get("parent_task_id") or None,
            "editable": editable,
            "is_milestone": bool(t.get("is_milestone")),
            "is_parent": is_parent,
            "include_in_dev_summary": bool(t.get("include_in_dev_summary")),
        })
        for c in sorted(children.get(t["id"], []), key=_sortkey):
            _emit(c, level + 1, member_matched)

    for r in sorted(visible_roots, key=_sortkey):
        _emit(r, 0)

    # 人物選択（上部プルダウン）：権限者（管理職・所属長・マスタ）は「全員」＋
    # 実効所属（assign_users、マスタは全ユーザー）から、表示・新規タスクの担当者
    # 既定値を選べるようにする。「全員」を選ぶとスコープ内全員分が混在表示される。
    people = ([{"id": 0, "name": "全員"}] + assign_users) if privileged else assign_users

    lu = get_user_by_id(login_user_id) or {}
    login_user = {"id": login_user_id,
                  "name": _display_name(lu.get("name", ""), lu.get("last_name"))}
    # 新規タスクの既定担当者：マスタ・管理職・所属長が特定ユーザーを選んでいればその人、
    # それ以外は操作者本人。
    if privileged and target_user_id not in (0, login_user_id):
        u = get_user_by_id(target_user_id) or {}
        default_assignee = {"id": target_user_id,
                            "name": _display_name(u.get("name", ""), u.get("last_name"))}
        view_name = default_assignee["name"]
    else:
        default_assignee = login_user
        view_name = "全員" if (privileged and target_user_id == 0) else login_user["name"]

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
        show_selector=privileged,
        can_parent=privileged,
        role=login_role,
        is_master=master,
        display_days=_DISPLAY_DAYS,
        holiday_dates=sorted(holiday_dates),
        csrf_token=session.get("csrf_token", ""),
        show_done=show_done,
        dev_summary_enabled=dev_summary_enabled,
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
    login_dept = session.get("user_dept", "")
    target_user_id = _resolve_target(login_user_id)
    show_done = request.args.get("show_done", "").strip() == "1"

    raw_start = request.args.get("start", "").strip()
    try:
        range_start = date.fromisoformat(raw_start) if raw_start else _default_range_start()
    except ValueError:
        range_start = _default_range_start()

    visible_roots, children = _get_visible_tree(login_user_id, login_role, target_user_id, login_dept)
    # Excel出力は4週間（28日）分のみ出力する（画面表示の9週とは別）。
    wb = _build_gantt_excel_by_tree(
        visible_roots, children, range_start, _EXPORT_DAYS,
        login_dept=login_dept, show_completed=show_done,
    )

    import io
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True, download_name=f"{_export_file_stem(login_dept)}.xlsx",
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
    login_dept = session.get("user_dept", "")
    target_user_id = _resolve_target(login_user_id)
    show_done = request.args.get("show_done", "").strip() == "1"

    raw_start = request.args.get("start", "").strip()
    try:
        range_start = date.fromisoformat(raw_start) if raw_start else _default_range_start()
    except ValueError:
        range_start = _default_range_start()

    visible_roots, children = _get_visible_tree(login_user_id, login_role, target_user_id, login_dept)
    # PDFもExcelと同じ4週間（28日）分を出力する。
    buf = _build_gantt_pdf_by_tree(
        visible_roots, children, range_start, _EXPORT_DAYS,
        login_dept=login_dept, show_completed=show_done,
    )

    return send_file(
        buf, as_attachment=True, download_name=f"{_export_file_stem(login_dept)}.pdf",
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

    # 役職による権限：親（レベル0）の作成・編集は管理職以上のみ。
    privileged = is_privileged(login_role)

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
                a1, a2, mem_str, assigned_ids_str = None, None, ",".join(str(i) for i in member_ids), ""
            else:
                # 子：担当者（2名以上対応）。未指定は既定担当者。
                a_ids = _parse_people(row) or [default_owner]
                a1 = a_ids[0]
                a2 = a_ids[1] if len(a_ids) > 1 else None
                mem_str = ""
                assigned_ids_str = ",".join(str(i) for i in a_ids)
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
                assigned_ids=assigned_ids_str,
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
                # 変更なしの行でも、単独タスクが子の追加によって事後的に親化した
                # ケースでは、残存する assigned_to をここでクリアする（このタスク
                # 自体は dirty ではないため update_project_task はフルコールしない）。
                if is_parent and (existing.get("assigned_to") or existing.get("assigned_to_2")):
                    if _can_touch_existing(existing):
                        clear_project_task_assigned_to(tid)
                continue
            if not _can_touch_existing(existing):
                continue
            assigned_ids_update: str | None = None
            if is_parent:
                # 親：メンバーを更新。親は担当者を持たない仕様のため常にクリアする
                # （過去のデータ不整合で残存した assigned_to が、子の全削除時に
                #   「担当者が復活する」ように見える不具合を防ぐ）。
                set_project_task_members(tid, ",".join(str(i) for i in _parse_people(row)))
                a1 = None
                a2 = None
            else:
                a_ids = _parse_people(row)
                a1 = a_ids[0] if a_ids else existing.get("assigned_to")
                a2 = (a_ids[1] if len(a_ids) > 1 else None) if a_ids else existing.get("assigned_to_2")
                assigned_ids_update = ",".join(str(i) for i in a_ids) if a_ids else ""
            if (s and e) or (not s and not e):
                # 開始日・終了日が両方揃っている場合は更新。両方空はガントバー削除（日程クリア）として扱う。
                update_project_task(
                    task_id=tid,
                    category_id=existing.get("category_id"),
                    subcategory_id=existing.get("subcategory_id"),
                    task_name=name or existing.get("task_name", ""),
                    description=existing.get("description", "") or "",
                    start_date=s or "", end_date=e or "",
                    status=_parse_status(row.get("status")),
                    progress=_parse_int(row.get("progress"), 0, 100),
                    delay_days=_parse_int(row.get("delay"), 0),
                    updated_by=session.get("user_name", ""),
                    assigned_to=a1,
                    assigned_to_2=a2,
                    is_milestone=existing.get("is_milestone", 0) or 0,
                    is_event=existing.get("is_event", 0) or 0,
                    planned_hours=existing.get("planned_hours", 0.0) or 0.0,
                    assigned_ids=assigned_ids_update,
                    include_in_dev_summary=(int(bool(row.get("dev_summary"))) if is_parent else None),
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
