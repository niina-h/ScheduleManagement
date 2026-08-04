"""管理職日報メール画面・設定ルート。"""
from __future__ import annotations

import html as html_mod
import urllib.parse
from datetime import date, timedelta
from email.mime.text import MIMEText

from flask import Blueprint, Response, abort, redirect, render_template, request, session, url_for

from ..auth_helpers import is_privileged, is_master
from ..models import (
    get_accessible_users,
    get_all_project_tasks,
    get_all_users,
    get_daily_result,
    get_daily_comment,
    get_events_for_user_date,
    get_global_task_category_map,
    get_mail_setting,
    get_next_business_day,
    get_task_master,
    get_weekly_leave,
    get_weekly_schedule,
    resolve_parent_progress,
    resolve_parent_status,
    save_mail_setting,
    get_user_by_id,
)

mail_report_bp = Blueprint("mail_report_bp", __name__, url_prefix="/mail-report")


def _require_privileged() -> None | object:
    """管理職以上（管理職・所属長・システム管理者）でなければリダイレクト／403を返す。

    Returns:
        None: 権限チェック通過時
        redirect / abort: 権限不足時
    """
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))
    if not is_privileged(session.get("user_role", "")):
        abort(403)
    return None


def _get_master_mail_members(login_user: dict) -> list[dict]:
    """マスタ用日報メールの対象メンバー（ログインユーザーの所属全員）を返す。

    get_accessible_users はシステム管理者の所属切替（session["active_dept"]）状態に
    連動するため、切替中は一部所属しか対象にならない。マスタ用メールは常に
    ログインユーザー自身の所属全体を対象とすべき機能のため、切替状態に関わらず
    dept 列で直接絞り込む。

    Args:
        login_user: ログインユーザー情報（dept を含む）。

    Returns:
        list[dict]: 所属メンバー一覧（ログインユーザー自身を含む）。
    """
    dept = login_user.get("dept", "")
    if not dept:
        return [login_user]
    return get_all_users(dept_filter=dept)


_WEEKDAY_JA = ["月", "火", "水", "木", "金", "土", "日"]


def _build_mgr_self_body(login_user: dict, target_date: date) -> tuple[str, str]:
    """管理職の自己日報メール件名・本文を生成する。

    Args:
        login_user: ログインユーザー情報
        target_date: 対象日

    Returns:
        tuple[str, str]: (件名, 本文)
    """
    date_str = target_date.isoformat()
    uid: int = login_user["id"]

    # 件名: mm/dd業務報告
    mm = f"{target_date.month:02d}"
    dd = f"{target_date.day:02d}"
    subject = f"{mm}/{dd}業務報告"

    # 振り返り・朝礼での気づき
    comment_row = get_daily_comment(uid, date_str)
    reflection = comment_row.get("reflection", "").strip() or "（未入力）"
    action = comment_row.get("action", "").strip() or "（未入力）"
    # 実施内容（本日の作業実績 - 同一作業名は1つにまとめる）
    result = get_daily_result(uid, date_str)
    work_seen: set[str] = set()
    work_order: list[str] = []
    for slot in ("am", "pm"):
        for item in result.get(slot, []):
            task = item.get("task_name", "").strip()
            if task and task not in work_seen:
                work_seen.add(task)
                work_order.append(task)
    work_results = "\n".join(work_order) if work_order else "（実績なし）"

    # 翌営業日の予定（同一作業名は1つにまとめる。
    # 土日・会社休日・本人の休暇設定をスキップして、実際に出勤する日の予定を表示する）
    next_day = get_next_business_day(uid, target_date)
    next_dow = next_day.weekday()  # 0=月〜4=金
    next_week_start = (next_day - timedelta(days=next_dow)).isoformat()
    next_schedule_data = get_weekly_schedule(uid, next_week_start)
    next_seen: list[str] = []
    for slot in ("am", "pm"):
        for item in next_schedule_data.get(next_dow, {}).get(slot, []):
            task = item.get("task_name", "").strip()
            if task and task not in next_seen:
                next_seen.append(task)
    next_schedule = "\n".join(next_seen) if next_seen else "（予定未入力）"

    body = (
        "お疲れ様です。\n"
        "本日の業務報告をいたします。\n"
        "\n"
        "＜本日の振り返り＞\n"
        f"{reflection}\n"
        "\n"
        "＜朝礼での気づき＞\n"
        f"{action}\n"
        "\n"
        "＜実施内容＞\n"
        f"{work_results}\n"
        "\n"
        "＜翌稼働日の達成目標＞\n"
        f"{next_schedule}\n"
        "\n"
        "以上になります。\n"
        "ご確認のほど、よろしくお願いいたします。\n"
        "\n"
        f"{login_user.get('name', '')}"
    )

    return subject, body


def _build_friday_report_default(login_user: dict, target_date: date) -> str:
    """金曜日用「管理業務のご報告」のデフォルトテキストを生成する。

    マスタ権限ユーザーの月〜金の実績を大区分ごとに集計し、
    (管理)(週間)(随時) の3カテゴリに分類して表示する。

    Args:
        login_user: ログインユーザー情報。
        target_date: 対象日（金曜日）。

    Returns:
        str: デフォルトの管理業務報告テキスト。
    """
    uid: int = login_user["id"]
    # 月〜金の日付を算出
    dow = target_date.weekday()  # 金曜=4
    monday = target_date - timedelta(days=dow)
    week_dates = [(monday + timedelta(days=d)).isoformat() for d in range(5)]

    # タスク→大区分マップ
    global_cat_map = get_global_task_category_map()
    task_master_list = get_task_master(uid)
    task_cat_map: dict[str, dict] = dict(global_cat_map)
    task_cat_map.update({
        t["task_name"]: {
            "category_name": t.get("category_name") or "",
            "subcategory_name": t.get("subcategory_name") or "",
        }
        for t in task_master_list if t.get("task_name")
    })

    # 大区分グループ別の累計時間を集計
    kanri_hours: float = 0.0   # 管理 + 事務
    weekly_hours: float = 0.0  # 開発 + ITインフラ
    zuiji_hours: float = 0.0   # サポート
    total_hours: float = 0.0

    for d_str in week_dates:
        result = get_daily_result(uid, d_str)
        for slot in ("am", "pm"):
            for item in result.get(slot, []):
                task = item.get("task_name", "").strip()
                hours = float(item.get("hours", 0.0))
                if not task or hours <= 0:
                    continue
                total_hours += hours
                cat = task_cat_map.get(task, {}).get("category_name", "")
                if cat in ("管理", "事務", "定例"):
                    kanri_hours += hours
                elif cat in ("開発", "ITインフラ"):
                    weekly_hours += hours
                elif cat in ("サポート",):
                    zuiji_hours += hours
                else:
                    kanri_hours += hours  # 未分類は管理に含む

    lines: list[str] = [
        "【管理業務のご報告】",
        "",
        "",
        "",
        "",
        "",
        f"　(管理）教育・進捗・事務　　　{kanri_hours:g}ｈ",
        f"　(週間）開発・AI・インフラ対応　{weekly_hours:g}ｈ",
        f"　(随時）問合せ対応　　　　　　　{zuiji_hours:g}ｈ",
    ]
    return "\n".join(lines)


def _get_friday_report(login_user: dict | None = None, target_date: date | None = None) -> str:
    """金曜日用「管理業務のご報告」テキストを取得する。

    保存済みテキストがあればそれを返し、なければデフォルトを生成する。

    Args:
        login_user: ログインユーザー情報（デフォルト生成用）。
        target_date: 対象日（デフォルト生成用）。

    Returns:
        str: 管理業務報告テキスト。
    """
    setting = get_mail_setting("マスタ_週次管理報告")
    saved = setting.get("body_template", "").strip()
    if saved:
        return saved
    if login_user and target_date:
        return _build_friday_report_default(login_user, target_date)
    return ""


def _save_friday_report(text: str) -> None:
    """金曜日用「管理業務のご報告」テキストを保存する。

    Args:
        text: 管理業務報告テキスト。
    """
    current = get_mail_setting("マスタ_週次管理報告")
    save_mail_setting(
        role="マスタ_週次管理報告",
        to_address=current.get("to_address", ""),
        cc_address=current.get("cc_address", ""),
        subject_template=current.get("subject_template", ""),
        body_template=text,
        bcc_address=current.get("bcc_address", ""),
    )


def _get_mgr_remarks() -> str:
    """管理職日報メールの備考欄テキストを取得する（印刷専用）。

    Returns:
        str: 備考テキスト（未設定なら空文字列）。
    """
    setting = get_mail_setting("管理職_備考")
    return setting.get("body_template", "").strip()


def _save_mgr_remarks(text: str) -> None:
    """管理職日報メールの備考欄テキストを保存する。

    Args:
        text: 備考テキスト。
    """
    current = get_mail_setting("管理職_備考")
    save_mail_setting(
        role="管理職_備考",
        to_address=current.get("to_address", ""),
        cc_address=current.get("cc_address", ""),
        subject_template=current.get("subject_template", ""),
        body_template=text,
        bcc_address=current.get("bcc_address", ""),
    )


def _build_master_subject(dept: str, target_date: date) -> str:
    """マスタ用メール件名を生成する（金曜は「管理・」付き）。

    Args:
        dept: 部署名
        target_date: 対象日

    Returns:
        str: メール件名
    """
    mm = f"{target_date.month:02d}"
    dd = f"{target_date.day:02d}"
    yyyy = str(target_date.year)
    dow = _WEEKDAY_JA[target_date.weekday()]
    if target_date.weekday() == 4:  # 金曜
        return f"【{dept}】管理・日次業務報告{yyyy}/{mm}/{dd}（{dow}）"
    return f"【{dept}】日次業務報告{yyyy}/{mm}/{dd}（{dow}）"


def _resolve_master_dept(login_user: dict, members: list[dict]) -> str:
    """マスタ用件名に表示する部署名を決定する。

    ログイン中マスタ自身の所属部署を最優先で使う。未設定の場合は、
    報告対象メンバーの中で最も多い部署名で補完し、件名の【】が
    空になるのを防ぐ。

    Args:
        login_user: ログインユーザー情報。
        members: 報告対象メンバーのリスト。

    Returns:
        str: 件名に表示する部署名（最終的に不明なら空文字）。
    """
    dept = (login_user.get("dept") or "").strip()
    if dept:
        return dept
    # フォールバック: メンバーの中で最も多い部署名を採用
    counts: dict[str, int] = {}
    for m in members:
        d = (m.get("dept") or "").strip()
        if d:
            counts[d] = counts.get(d, 0) + 1
    if counts:
        return max(counts, key=lambda k: counts[k])
    return ""


def _build_master_body(
    login_user: dict, target_date: date, members: list[dict], greeting: str,
    friday_report: str = "",
) -> str:
    """マスタ用メール本文を動的生成する。

    大区分・中区分でグループ化した作業実績と、メンバー別AM/PMサマリ、
    振り返り、AI開発状況、次回予定を含む。
    金曜日の場合は「管理業務のご報告」セクションを挨拶文の直後に挿入する。

    Args:
        login_user: ログインユーザー情報
        target_date: 対象日
        members: アクセス可能なメンバー一覧
        greeting: 宛先挨拶文（設定から取得）
        friday_report: 金曜日用「管理業務のご報告」テキスト（空なら挿入しない）

    Returns:
        str: メール本文
    """
    date_str = target_date.isoformat()
    dept = login_user.get("dept", "")
    login_id: int = login_user["id"]

    # 対象日の曜日・週開始日（週次スケジュール取得用）
    target_dow = target_date.weekday()
    week_start_str = (target_date - timedelta(days=target_dow)).isoformat()

    # 全ユーザー横断の区分マップ（フォールバック用）
    global_cat_map = get_global_task_category_map()

    # タスク名 → project_task.id の逆引き（週間予定に project_task_id が
    # 未設定のまま残っている行を、タスク名の完全一致で補完するために使う）。
    # 本番データでタスク名の重複が無いことを確認済み（重複時は先に見つかった方を使う）。
    task_name_to_id: dict[str, int] = {}
    for t in get_all_project_tasks():
        task_name_to_id.setdefault(t["task_name"], t["id"])

    # 各メンバーの実績・タスクマスタ・週次スケジュールを収集
    member_data: list[dict] = []
    for member in members:
        uid = member["id"]
        result = get_daily_result(uid, date_str)
        comment_row = get_daily_comment(uid, date_str)
        task_master_list = get_task_master(uid)
        # task_name → {category_name, subcategory_name} マップ（個人＋グローバルフォールバック）
        task_cat_map: dict[str, dict] = dict(global_cat_map)
        task_cat_map.update({
            t["task_name"]: {
                "category_name": t.get("category_name") or "",
                "subcategory_name": t.get("subcategory_name") or "",
            }
            for t in task_master_list
        })
        # 対象日の週次スケジュール（計画判定用）
        schedule_data = get_weekly_schedule(uid, week_start_str)
        today_schedule_items = [
            item
            for slot in ("am", "pm")
            for item in schedule_data.get(target_dow, {}).get(slot, [])
            if item.get("task_name", "").strip()
        ]
        scheduled_tasks: set[str] = {
            item["task_name"].strip() for item in today_schedule_items
        }
        # 本日の週間予定から project_task の ID を解決する（日報メール「対応中」
        # セクションの判定に使用。実績入力の有無は問わない）。
        # project_task_id が未設定の行は、タスク名の完全一致で project_task を
        # 逆引きして補完する（ガント反映後の手直しなどで紐付けが欠けているケースが
        # 多数あるため）。
        scheduled_task_ids: set[int] = set()
        for item in today_schedule_items:
            pt_id = item.get("project_task_id")
            if not pt_id:
                pt_id = task_name_to_id.get(item["task_name"].strip())
            if pt_id:
                scheduled_task_ids.add(pt_id)
        # 当日の休暇種別（1日有休 / AM半休 / PM半休 / 特休 / 祝日 / その他休み）
        leave_type: str = get_weekly_leave(uid, week_start_str).get(target_dow, "")
        member_data.append({
            "member": member,
            "result": result,
            "comment": comment_row,
            "task_cat_map": task_cat_map,
            "scheduled_tasks": scheduled_tasks,
            "scheduled_task_ids": scheduled_task_ids,
            "leave_type": leave_type,
        })

    # 全メンバーの予定時間合計（週間スケジュールの入力時間）
    total_planned_hours = 0.0
    for md in member_data:
        schedule_data = get_weekly_schedule(md["member"]["id"], week_start_str)
        for slot in ("am", "pm"):
            for item in schedule_data.get(target_dow, {}).get(slot, []):
                h = float(item.get("hours", 0.0))
                if item.get("task_name", "").strip() and h > 0:
                    total_planned_hours += h

    # 週間予定の入力率（休暇者は対象人数から除外）：
    # 対象日にタスク名付きの予定が1件以上あるメンバー数 / 休暇者を除いた全体人数
    schedule_targets = [md for md in member_data if not md["leave_type"]]
    schedule_entered = sum(1 for md in schedule_targets if md["scheduled_tasks"])
    if schedule_targets:
        schedule_entry_rate = round(schedule_entered / len(schedule_targets) * 100)
    else:
        schedule_entry_rate = 0

    # 計画/突発/リスケ の時間集計
    plan_hours = 0.0
    sudden_hours = 0.0
    resc_hours = 0.0
    total_actual_hours = 0.0
    for md in member_data:
        for slot in ("am", "pm"):
            for item in md["result"].get(slot, []):
                task = item.get("task_name", "").strip()
                hours = float(item.get("hours", 0.0))
                if not task or hours == 0.0:
                    continue
                total_actual_hours += hours
                if int(item.get("is_carryover", 0)):
                    resc_hours += hours
                elif task in md["scheduled_tasks"]:
                    plan_hours += hours
                else:
                    sudden_hours += hours

    # 予定時間ベース（実績% = 計画% + 突発% + リスケ%）
    if total_planned_hours > 0:
        plan_rate = round(plan_hours / total_planned_hours * 100)
        sudden_rate = round(sudden_hours / total_planned_hours * 100)
        resc_rate = round(resc_hours / total_planned_hours * 100)
        jisseki_rate = plan_rate + sudden_rate + resc_rate
    else:
        plan_rate = sudden_rate = resc_rate = jisseki_rate = 0

    # 業務内容セクション（project_task の登録タスクを表示）
    # タスク一覧画面と同じスコープ（担当メンバーのタスクのみ、イベント除外）。
    member_ids = [m["id"] for m in members]
    target_date_str: str = target_date.isoformat()
    project_tasks = [
        t for t in get_all_project_tasks(user_ids=member_ids)
        if not t.get("is_event", 0)
    ]

    # 「対応中」判定: 週間予定・日次実績のタスク名一致には一切依存せず、
    # ガントチャートの登録内容から対象日時点の状況を判定する（実績・予定の
    # 入力有無に関わらず、所属メンバー全員のガント登録が反映される）。
    # 子・単独タスク（担当者を持つ末端）が対象。親（見出し行、バー無し）は子の
    # 状況で判定する（親自身に start_date/end_date が無いことが多いため）。
    # 子は絞り込み前の全タスクから解決する（子の担当者がメール対象メンバーと
    # 異なる場合も正しく集計するため）。
    _all_tasks_for_children = get_all_project_tasks()
    all_tasks_by_id: dict[int, dict] = {t["id"]: t for t in _all_tasks_for_children}
    _children_by_parent: dict[int, list[dict]] = {}
    for _t in _all_tasks_for_children:
        _pid = _t.get("parent_task_id")
        if _pid:
            _children_by_parent.setdefault(_pid, []).append(_t)

    def _in_progress_on_target_date(t: dict) -> bool:
        """このタスク（子・単独）が対象日時点で「対応中」に該当するか判定する。

        未完了タスクは開始日〜終了日に対象日が含まれるかで判定する。完了タスクは
        「完了にした日」を実態の作業日とみなす：終了日以前に完了していれば
        updated_at（完了操作日）が対象日と一致するか、終了日より後にずれ込んで
        完了していれば終了日が対象日と一致するかで判定する。
        """
        start = (t.get("start_date") or "").strip()
        end = (t.get("end_date") or "").strip()
        if not start or not end:
            return False
        if t.get("status") != "完了":
            return start <= target_date_str <= end
        updated_date = (t.get("updated_at") or "").strip()[:10]
        if updated_date and updated_date <= end:
            return updated_date == target_date_str
        return end == target_date_str

    in_progress_tasks = [
        t for t in project_tasks
        if _in_progress_on_target_date(t)
        or any(
            _in_progress_on_target_date(c)
            for c in _children_by_parent.get(t["id"], [])
        )
    ]

    # 親タスク（バー無しの見出し行）は、ガントチャート・進捗ダッシュボードと表示を
    # 一致させるため、子タスクの状況から状態・進捗を自動判定する（DBの生値のまま
    # だと常に「未着手・0%」に見えてしまう）。
    for pt in in_progress_tasks:
        _kids = _children_by_parent.get(pt["id"])
        if _kids:
            pt["status"] = resolve_parent_status(pt, _kids)
            pt["progress"] = resolve_parent_progress(pt, _kids)

    all_users_by_id: dict[int, dict] = {u["id"]: u for u in get_all_users()}

    def _assignee_names(pt: dict) -> str:
        """タスクの担当者名（親ならメンバー、子なら担当者1・2）を「、」区切りで返す（姓のみ）。"""
        member_ids_raw = (pt.get("member_ids") or "").strip()
        if member_ids_raw:
            names = [
                (all_users_by_id[int(s)].get("last_name") or all_users_by_id[int(s)]["name"])
                for s in member_ids_raw.split(",")
                if s.strip().isdigit() and int(s) in all_users_by_id
            ]
            return "、".join(names) if names else "担当者未設定"
        names = []
        if pt.get("assigned_name"):
            names.append(pt.get("assigned_last_name") or pt.get("assigned_name"))
        if pt.get("assigned_name_2"):
            names.append(pt.get("assigned_last_name_2") or pt.get("assigned_name_2"))
        return "、".join(names) if names else "担当者未設定"

    def _effective_category(pt: dict) -> dict:
        """区分表示用の情報（中区分名・大区分名・並び順）を返す。

        タスク自身に区分が未設定の場合、親タスクの区分をフォールバックとして使う
        （「その他」への集約を減らすため。子は業務内容までは登録するが区分登録が
        後回しになりがちで、親には区分が設定済みのケースが多いことが判明したため）。
        """
        if pt.get("subcategory_name") or pt.get("category_name"):
            return pt
        parent_id = pt.get("parent_task_id")
        parent = all_tasks_by_id.get(parent_id) if parent_id else None
        if parent and (parent.get("subcategory_name") or parent.get("category_name")):
            return parent
        return pt

    def _root_ancestor(pt: dict) -> dict:
        """タスクの最上位（ルート）の祖先タスクを返す（親を持たなければ自分自身）。

        ガントチャートは3階層以上の親子関係を持つことがあり（例：配筋システム開発→
        Revit運用環境開発設計→Revit連携システム開発）、メールの見出しはガント画面
        最上段の項目名と一致させる必要があるため、1階層だけでなくルートまで辿る。
        循環参照があっても無限ループしないよう、訪問済みIDを記録して打ち切る。
        """
        seen: set[int] = set()
        cur = pt
        while True:
            parent_id = cur.get("parent_task_id")
            if not parent_id or parent_id in seen:
                return cur
            parent = all_tasks_by_id.get(parent_id)
            if not parent:
                return cur
            seen.add(parent_id)
            cur = parent

    def _group_label(pt: dict) -> str:
        """「対応中」の見出しラベルを返す（ガントチャート最上位の親タスク名を使う）。

        単独タスク（親を持たない）は自分自身のタスク名を見出しとする（1件だけの
        見出しになる）。タスク区分マスタ（大区分・中区分）には依存しない——区分の
        登録有無に関わらずガントチャートの見た目と一致させる。
        """
        return _root_ancestor(pt)["task_name"]

    # ガントチャート上の親タスク名ごとにグループ化する（区分マスタの登録有無に
    # 関わらず、ガントチャートで見えている親子構造とメールの見出しを一致させる）。
    tasks_by_subcat: dict[str, list[dict]] = {}
    subcat_order: dict[str, int] = {}
    for pt in in_progress_tasks:
        label = _group_label(pt)
        tasks_by_subcat.setdefault(label, []).append(pt)
        cat_src = _effective_category(pt)
        order = cat_src.get("subcat_order") or cat_src.get("cat_order") or 0
        if label not in subcat_order or order < subcat_order[label]:
            subcat_order[label] = order

    total_count = sum(len(v) for v in tasks_by_subcat.values())

    content_lines: list[str] = []
    content_lines.append(f"対応中 全{total_count}件")
    for label in sorted(tasks_by_subcat.keys(), key=lambda k: (subcat_order.get(k, 0), k)):
        group = sorted(tasks_by_subcat[label], key=lambda t: t.get("display_order") or 0)
        # 単独タスク（親を持たず、見出し自体がそのタスク名）の場合は、見出し行に
        # 直接進捗情報を付記し、名前が重複する子行は出さない。
        if len(group) == 1 and group[0]["task_name"] == label:
            pt = group[0]
            progress = pt.get("progress", 0) or 0
            status = pt.get("status", "") or "未着手"
            if progress >= 1:
                content_lines.append(f"　・{label}　（{_assignee_names(pt)}：進捗{progress}%、{status}）")
            else:
                content_lines.append(f"　・{label}")
            continue
        content_lines.append(f"　・{label}")
        for pt in group:
            progress = pt.get("progress", 0) or 0
            status = pt.get("status", "") or "未着手"
            if progress >= 1:
                content_lines.append(f"　　└{pt['task_name']}　（{_assignee_names(pt)}：進捗{progress}%、{status}）")
            else:
                content_lines.append(f"　　└{pt['task_name']}")

    # メンバー実績サマリ
    # 除外条件: 定例作業（大区分「定例」or 中区分「定例作業」）、AM1行目(idx=0)、PM最終行(idx=4)
    def _is_routine(task_name: str, cat_map: dict) -> bool:
        """定例作業かどうか判定する。"""
        info = cat_map.get(task_name.strip(), {})
        return (info.get("category_name") in ("定例", "定例作業")
                or info.get("subcategory_name") in ("定例", "定例作業"))

    def _parent_task_name(project_task_id: int | None) -> str | None:
        """project_task_id から親タスク名を返す（親を持たない・紐付けなしなら None）。"""
        if not project_task_id:
            return None
        task = all_tasks_by_id.get(project_task_id)
        if not task:
            return None
        parent_id = task.get("parent_task_id")
        parent = all_tasks_by_id.get(parent_id) if parent_id else None
        return parent["task_name"] if parent else None

    def _format_member_tasks(items: list[dict]) -> str:
        """作業名の後ろに親タスク名を付けて連結する。連続する項目の親が同じ場合は省略する。"""
        parts: list[str] = []
        last_parent: str | None = "__init__"  # 最初の項目は必ず親名を出すための番兵
        for name, parent_name in items:
            if parent_name and parent_name != last_parent:
                parts.append(f"{name}（{parent_name}）")
            else:
                parts.append(name)
            last_parent = parent_name
        return "/ ".join(parts) if parts else "（なし）"

    # 終日休暇（出勤なし）と判定する種別
    FULL_LEAVE_TYPES: set[str] = {"1日有休", "特休", "その他休み", "祝日"}

    member_lines: list[str] = []
    for md in member_data:
        name = md["member"].get("last_name") or md["member"]["name"]
        result = md["result"]
        tcm = md["task_cat_map"]
        leave = md.get("leave_type", "")

        # 終日休暇の場合：休暇種別のみを表示
        if leave in FULL_LEAVE_TYPES:
            member_lines.append(f"{name}：{leave}")
            continue

        # AM・PMを通しで1本にまとめる（AM1行目・PM最終行・定例作業は除外）。
        seen_names: set[str] = set()
        combined_items: list[tuple[str, str | None]] = []
        for idx, item in enumerate(result.get("am", [])):
            task_name = item.get("task_name", "").strip()
            if (not task_name or float(item.get("hours", 0)) <= 0
                    or idx == 0 or _is_routine(task_name, tcm) or task_name in seen_names):
                continue
            seen_names.add(task_name)
            combined_items.append((task_name, _parent_task_name(item.get("project_task_id"))))
        for idx, item in enumerate(result.get("pm", [])):
            task_name = item.get("task_name", "").strip()
            if (not task_name or float(item.get("hours", 0)) <= 0
                    or idx == 4 or _is_routine(task_name, tcm) or task_name in seen_names):
                continue
            seen_names.add(task_name)
            combined_items.append((task_name, _parent_task_name(item.get("project_task_id"))))

        tasks_str = _format_member_tasks(combined_items)

        # 半休の場合は先頭に休暇種別を付記する。
        if leave == "AM半休":
            member_lines.append(f"{name}：AM半休 / {tasks_str}")
        elif leave == "PM半休":
            member_lines.append(f"{name}：{tasks_str} / PM半休")
        else:
            member_lines.append(f"{name}：{tasks_str}")

    # マスタ自身の振り返り・対策
    def _wrap_text(text: str, width: int = 100) -> str:
        """テキストを指定幅で改行する。"""
        if len(text) <= width:
            return text
        lines: list[str] = []
        while len(text) > width:
            # 幅以内の最後の句読点・カンマで改行
            pos = -1
            for ch in ("。", "、", "，", ".", ",", "　"):
                p = text.rfind(ch, 0, width)
                if p > pos:
                    pos = p
            if pos <= 0:
                pos = width  # 句読点なければ強制改行
            lines.append(text[:pos + 1])
            text = text[pos + 1:]
        if text:
            lines.append(text)
        return "\n".join(lines)

    master_comment = get_daily_comment(login_id, date_str)
    reflection = _wrap_text(master_comment.get("reflection", "").strip() or "（未入力）")

    # ＜開発状況＞: 大区分「開発」の project_task を対象に、誰が何を対応しているかを
    # 進捗率・状態とあわせて表示する。
    # （日次実績の自由入力タスク名はタスクマスタに区分登録されていないことが多く、
    #   実績ベースの集計では常に空になりがちなため、大区分が確実に設定されている
    #   project_task を情報源にする。）
    #   形式: 「  中区分　タスク名　（姓：進捗X%、状態）」
    # 担当者名は _assignee_names（姓のみ表示、上の「対応中」セクションで定義済み）を再利用する。
    dev_tasks = [
        pt for pt in project_tasks
        if (_effective_category(pt).get("category_name") or "") == "開発"
        and _in_progress_on_target_date(pt)
    ]
    # 中区分ごとにグループ化する。進捗1%以上のタスクのみ担当者・進捗・状態を付記し、
    # それ以外（未着手）はタスク名のみを列記して情報量を抑える。
    # 並び順はガントチャートと同じ表示順（subcategory の display_order）で統一する
    # （現状は開発の中区分が全て display_order=0 のため実質アルファベット順になって
    #   いたので、rows挿入順にも依存しない安定した順序を明示する）。
    dev_by_subcat: dict[str, list[dict]] = {}
    dev_subcat_order: dict[str, int] = {}
    for pt in dev_tasks:
        cat_src = _effective_category(pt)
        label = (cat_src.get("subcategory_name") or "").strip() or "開発"
        dev_by_subcat.setdefault(label, []).append(pt)
        order = cat_src.get("subcat_order") or 0
        if label not in dev_subcat_order or order < dev_subcat_order[label]:
            dev_subcat_order[label] = order

    ai_lines: list[str] = []
    for label in sorted(dev_by_subcat.keys(), key=lambda k: (dev_subcat_order.get(k, 0), k)):
        ai_lines.append(f"  {label}")
        for pt in sorted(dev_by_subcat[label], key=lambda t: t.get("display_order") or 0):
            progress = pt.get("progress", 0) or 0
            if progress >= 1:
                assignees = _assignee_names(pt)
                status = pt.get("status", "") or "未着手"
                ai_lines.append(f"　　{pt['task_name']}　（{assignees}：進捗{progress}%、{status}）")
            else:
                ai_lines.append(f"　　{pt['task_name']}")
    ai_section = "\n".join(ai_lines) if ai_lines else "  （開発中のタスクなし）"

    # ＜次回予定＞: マスタ自身の翌営業日予定（定例作業は除外）
    # 土日・会社休日・本人の休暇設定をスキップして、実際に出勤する日の予定を表示する
    next_day = get_next_business_day(login_id, target_date)
    next_dow = next_day.weekday()
    next_week_start = (next_day - timedelta(days=next_dow)).isoformat()
    next_schedule_data = get_weekly_schedule(login_id, next_week_start)
    # マスタ自身の区分マップを取得
    master_md = next((md for md in member_data if md["member"]["id"] == login_id), None)
    master_tcm = master_md["task_cat_map"] if master_md else {}
    next_seen_list: list[str] = []
    for slot in ("am", "pm"):
        for item in next_schedule_data.get(next_dow, {}).get(slot, []):
            t = item.get("task_name", "").strip()
            if t and t not in next_seen_list and not _is_routine(t, master_tcm):
                next_seen_list.append(t)
    # 翌日のイベントを全メンバーから収集
    next_date_str = next_day.isoformat()
    next_events_seen: list[str] = []
    for md in member_data:
        uid = md["member"]["id"]
        events = get_events_for_user_date(uid, next_date_str)
        for ev in events:
            name = ev.get("task_name", "").strip()
            time_str = ""
            if ev.get("event_start_time") and ev.get("event_end_time"):
                time_str = f"（{ev['event_start_time']}〜{ev['event_end_time']}）"
            label = f"📅 {name}{time_str}"
            if label not in next_events_seen:
                next_events_seen.append(label)

    next_lines: list[str] = [f"・{t}" for t in next_seen_list]
    if next_events_seen:
        next_lines.extend(next_events_seen)
    next_schedule = "\n".join(next_lines) if next_lines else "（予定未入力）"

    # 本文組み立て
    parts: list[str] = []
    if greeting.strip():
        parts.append(greeting.strip())
        parts.append("")
    parts.append(f"お疲れ様です。{dept}の業務報告となります。")
    parts.append("")

    # 金曜日: 管理業務のご報告を挿入
    if target_date.weekday() == 4 and friday_report.strip():
        parts.append(friday_report.strip())
        parts.append("")

    parts.extend([
        f"□予定：{schedule_entry_rate}%（週間予定を入力済みのメンバー：{schedule_entered}/{len(schedule_targets)}人）",
        f"■実績：{jisseki_rate}%（計画：{plan_rate}%　突発：{sudden_rate}%　リスケ：{resc_rate}%）",
        "\n".join(content_lines),
        "",
        "\n".join(member_lines),
        "",
        "＜本日の振り返り＞",
        reflection,
        "",
        "＜開発状況＞",
        ai_section,
        "",
        "＜次回予定＞",
        next_schedule,
        "",
        "以上になります。ご確認のほど、よろしくお願いいたします。",
    ])
    return "\n".join(parts)


def _build_member_reports(login_user: dict, date_str: str) -> str:
    """メンバーの実績をテキスト形式で組み立てる。

    Args:
        login_user: ログインユーザー情報（id, role, dept, name を含む）
        date_str: 日付（'YYYY-MM-DD'形式）

    Returns:
        str: 各メンバーの実績テキスト（改行区切り）
    """
    login_id = login_user["id"]
    login_role = login_user["role"]
    login_dept = login_user.get("dept", "")
    members = get_accessible_users(login_id, login_role, login_dept)

    lines: list[str] = []
    for member in members:
        uid = member["id"]
        name = member["name"]
        result = get_daily_result(uid, date_str)
        comment_row = get_daily_comment(uid, date_str)
        comment = comment_row.get("reflection", "") if comment_row else ""

        lines.append(f"【{name}】")

        # AM・PM全スロットを収集（task_nameが空でないもの）
        entries: list[str] = []
        for slot in ("am", "pm"):
            slot_list = result.get(slot, [])
            for item in slot_list:
                task = item.get("task_name", "").strip()
                hours = item.get("hours", 0)
                if task:
                    entries.append(f"  {task}  {hours}h")

        if entries:
            lines.extend(entries)
        else:
            lines.append("  （実績なし）")

        if comment:
            lines.append(f"  コメント: {comment}")
        lines.append("")

    return "\n".join(lines)


def _build_mailto(setting: dict, subject: str, *, include_body: str = "") -> str:
    """mailto: URLを組み立てる。

    Args:
        setting: メール設定 dict（to_address, cc_address, bcc_address を含む）
        subject: メール件名
        include_body: 本文（空文字列の場合は本文なし＝署名が保持される）

    Returns:
        str: mailto: スキームのURL文字列
    """
    params: dict[str, str] = {"subject": subject}
    if include_body:
        params["body"] = include_body
    if setting.get("cc_address"):
        params["cc"] = setting["cc_address"]
    if setting.get("bcc_address"):
        params["bcc"] = setting["bcc_address"]
    query = urllib.parse.urlencode(params, quote_via=urllib.parse.quote)
    to = urllib.parse.quote(setting.get("to_address", ""))
    return f"mailto:{to}?{query}"


def _build_eml(setting: dict, subject: str, body: str) -> str:
    """HTML形式の.emlファイルコンテンツを生成する。

    フォントを游ゴシック 11ptで統一したHTMLメールを生成。
    Outlookで開いた際にフォントが統一される。

    Args:
        setting: メール設定 dict（to_address, cc_address, bcc_address を含む）
        subject: メール件名
        body: メール本文（プレーンテキスト）

    Returns:
        str: .emlファイルの文字列
    """
    escaped = html_mod.escape(body)
    html_body = (
        '<html><head><meta charset="utf-8"></head>'
        '<body style="font-family: \'游ゴシック\', \'Yu Gothic\', sans-serif; font-size: 11pt; line-height: 1.6;">'
        f'<pre style="font-family: \'游ゴシック\', \'Yu Gothic\', sans-serif; font-size: 11pt; '
        f'white-space: pre-wrap; margin: 0;">{escaped}</pre>'
        '</body></html>'
    )
    msg = MIMEText(html_body, "html", "utf-8")
    msg["Subject"] = subject
    msg["To"] = setting.get("to_address", "")
    if setting.get("cc_address"):
        msg["Cc"] = setting["cc_address"]
    if setting.get("bcc_address"):
        msg["Bcc"] = setting["bcc_address"]
    # Outlookで「編集可能な下書き」として開くためのマーカー（classic Outlookが解釈）
    msg["X-Unsent"] = "1"
    return msg.as_string()


@mail_report_bp.route("/download_eml")
def download_eml():
    """メール内容を.emlファイルとしてダウンロードする。

    Outlookで開くとHTMLメールとして表示され、フォントが統一される。

    Returns:
        Response: .emlファイルのダウンロードレスポンス
    """
    redir = _require_privileged()
    if redir:
        return redir

    raw_date = request.args.get("date", "").strip()
    role_type = request.args.get("type", "mgr")
    if role_type not in ("mgr", "master"):
        role_type = "mgr"
    try:
        target_date = date.fromisoformat(raw_date)
    except ValueError:
        target_date = date.today()

    login_user = get_user_by_id(int(session["user_id"]))
    if not login_user:
        abort(404)

    if role_type == "master":
        setting = get_mail_setting("マスタ")
        members = _get_master_mail_members(login_user)
        subject = _build_master_subject(_resolve_master_dept(login_user, members), target_date)
        greeting = setting.get("body_template", "")
        body = _build_master_body(login_user, target_date, members, greeting, _get_friday_report(login_user, target_date))
    else:
        setting = get_mail_setting("管理職")
        subject, body = _build_mgr_self_body(login_user, target_date)

    eml_content = _build_eml(setting, subject, body)
    filename = f"daily_report_{target_date.isoformat()}_{role_type}.eml"

    return Response(
        eml_content,
        mimetype="message/rfc822",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@mail_report_bp.route("/preview")
def preview():
    """管理職日報メールのプレビュー画面。"""
    redir = _require_privileged()
    if redir:
        return redir

    raw_date = request.args.get("date", "").strip()
    try:
        target_date = date.fromisoformat(raw_date)
        date_str = target_date.isoformat()
    except ValueError:
        target_date = date.today()
        date_str = target_date.isoformat()

    WEEKDAY_JA = ["月", "火", "水", "木", "金", "土", "日"]
    day_of_week = WEEKDAY_JA[target_date.weekday()]
    date_display = f"{target_date.year}/{target_date.month:02d}/{target_date.day:02d}"

    login_user = get_user_by_id(int(session["user_id"]))
    if not login_user:
        abort(404)

    login_role = session.get("user_role", "")
    uid = login_user["id"]
    mgr_setting = _get_mgr_setting(uid)
    master_setting = _get_master_setting(uid)

    # 管理職用: 固定テンプレート
    mgr_subject, mgr_body = _build_mgr_self_body(login_user, target_date)
    mgr_mailto = _build_mailto(mgr_setting, mgr_subject)

    # マスタ用: 動的生成（件名は曜日で自動判定、本文は大区分・中区分グループ化）
    members = _get_master_mail_members(login_user)
    master_subject = _build_master_subject(_resolve_master_dept(login_user, members), target_date)
    master_greeting = master_setting.get("body_template", "")
    master_body = _build_master_body(login_user, target_date, members, master_greeting, _get_friday_report(login_user, target_date))
    master_mailto = _build_mailto(master_setting, master_subject)

    # 金曜日判定・管理業務報告テキスト
    is_friday: bool = target_date.weekday() == 4
    friday_report: str = _get_friday_report(login_user, target_date) if is_friday else ""

    # 備考欄（印刷専用）
    mgr_remarks: str = _get_mgr_remarks()

    return render_template(
        "mail_report_preview.html",
        date_str=date_str,
        date_display=date_display,
        day_of_week=day_of_week,
        login_role=login_role,
        mgr_setting=mgr_setting,
        master_setting=master_setting,
        mgr_subject=mgr_subject,
        mgr_body=mgr_body,
        mgr_mailto=mgr_mailto,
        master_subject=master_subject,
        master_body=master_body,
        master_mailto=master_mailto,
        csrf_token=session.get("csrf_token", ""),
        is_friday=is_friday,
        friday_report=friday_report,
        mgr_remarks=mgr_remarks,
    )


@mail_report_bp.route("/print-master")
def print_master() -> object:
    """マスタ用メール本文の印刷専用ページを返す。

    A4 1ページに収まる最小限のHTMLを返し、ブラウザの印刷／PDF保存で使用する。

    Returns:
        str: 印刷用HTML
    """
    redir = _require_privileged()
    if redir:
        return redir

    raw_date = request.args.get("date", "").strip()
    try:
        target_date = date.fromisoformat(raw_date)
    except ValueError:
        target_date = date.today()

    login_user = get_user_by_id(int(session["user_id"]))
    if not login_user:
        abort(404)

    members = _get_master_mail_members(login_user)
    master_subject = _build_master_subject(_resolve_master_dept(login_user, members), target_date)
    master_greeting = _get_master_setting(login_user["id"]).get("body_template", "")
    master_body = _build_master_body(login_user, target_date, members, master_greeting, _get_friday_report(login_user, target_date))

    escaped_body = html_mod.escape(master_body)
    mgr_remarks = _get_mgr_remarks()

    return render_template(
        "mail_report_print.html",
        subject=master_subject,
        body=escaped_body,
        mgr_remarks=mgr_remarks,
    )


@mail_report_bp.route("/save-address", methods=["POST"])
def save_address() -> object:
    """管理職日報プレビュー画面からTO・CCを保存する。

    Returns:
        object: プレビュー画面へのリダイレクト
    """
    redir = _require_privileged()
    if redir:
        return redir

    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    role: str = request.form.get("role", "")
    if role not in ("管理職", "マスタ"):
        abort(400)

    uid = int(session["user_id"])
    setting_key = _mgr_setting_key(uid) if role == "管理職" else _master_setting_key(uid)
    current = get_mail_setting(setting_key)
    save_mail_setting(
        role=setting_key,
        to_address=request.form.get("to_address", "").strip(),
        cc_address=request.form.get("cc_address", "").strip(),
        subject_template=current.get("subject_template", ""),
        body_template=current.get("body_template", ""),
        bcc_address=request.form.get("bcc_address", "").strip(),
    )

    date_str = request.form.get("date_str", "")
    return redirect(url_for("mail_report_bp.preview", date=date_str))


@mail_report_bp.route("/save-friday-report", methods=["POST"])
def save_friday_report() -> object:
    """金曜日用「管理業務のご報告」テキストを保存する。

    Returns:
        object: プレビュー画面へのリダイレクト。
    """
    redir = _require_privileged()
    if redir:
        return redir
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    _save_friday_report(request.form.get("friday_report", ""))
    date_str = request.form.get("date_str", "")
    return redirect(url_for("mail_report_bp.preview", date=date_str))


@mail_report_bp.route("/save-mgr-remarks", methods=["POST"])
def save_mgr_remarks() -> object:
    """管理職日報メールの備考欄テキストを保存する（印刷専用）。

    Returns:
        object: プレビュー画面へのリダイレクト。
    """
    redir = _require_privileged()
    if redir:
        return redir
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    _save_mgr_remarks(request.form.get("mgr_remarks", ""))
    date_str = request.form.get("date_str", "")
    return redirect(url_for("mail_report_bp.preview", date=date_str))


@mail_report_bp.route("/settings", methods=["GET", "POST"])
def settings():
    """メール設定画面（マスタのみ）。"""
    redir = _require_privileged()
    if redir:
        return redir
    if not is_master(session.get("user_role", "")):
        abort(403)

    if request.method == "POST":
        form_csrf = request.form.get("csrf_token", "")
        session_csrf = session.get("csrf_token", "")
        if form_csrf != session_csrf:
            abort(400)

        for role in ("管理職", "マスタ"):
            prefix = "mgr" if role == "管理職" else "master"
            save_mail_setting(
                role=role,
                to_address=request.form.get(f"{prefix}_to", "").strip(),
                cc_address=request.form.get(f"{prefix}_cc", "").strip(),
                subject_template=request.form.get(f"{prefix}_subject", "").strip(),
                body_template=request.form.get(f"{prefix}_body", "").strip(),
                bcc_address=request.form.get(f"{prefix}_bcc", "").strip(),
            )
        return redirect(url_for("mail_report_bp.settings"))

    mgr_setting = get_mail_setting("管理職")
    master_setting = get_mail_setting("マスタ")

    return render_template(
        "mail_report_settings.html",
        mgr_setting=mgr_setting,
        master_setting=master_setting,
        csrf_token=session.get("csrf_token", ""),
    )


# ====================================================================
# ユーザー用 日報メール
# ====================================================================

_USER_DEFAULT_BODY = (
    "お疲れ様です。\n"
    "\n"
    "本日の作業内容について、作業報告書を送付いたします。\n"
    "\n"
    "よろしくお願いいたします。\n"
)


def _build_user_subject(user: dict, target_date: date) -> str:
    """ユーザー用メール件名を生成する。

    Args:
        user: ユーザー情報辞書（'last_name' を含む）。
        target_date: 対象日。

    Returns:
        str: 件名文字列。
    """
    WEEKDAY_JA = ["月", "火", "水", "木", "金", "土", "日"]
    dow = WEEKDAY_JA[target_date.weekday()]
    last_name = user.get("last_name", "") or user.get("name", "")
    return f'日次作業報告「{last_name}」：{target_date.year}/{target_date.month:02d}/{target_date.day:02d}（{dow}）'


def _mgr_setting_key(user_id: int) -> str:
    """管理職用タブのメール設定キー（ユーザー個別）を返す。"""
    return f"管理職_{user_id}"


def _master_setting_key(user_id: int) -> str:
    """マスタ用タブのメール設定キー（ユーザー個別）を返す。"""
    return f"マスタ_{user_id}"


def _is_empty_setting(setting: dict) -> bool:
    """メール設定が未保存（全項目空）かどうかを判定する。"""
    return not any(
        (setting.get(k) or "").strip()
        for k in ("to_address", "cc_address", "bcc_address", "subject_template", "body_template")
    )


def _get_mgr_setting(user_id: int) -> dict:
    """管理職用タブのメール設定を取得する（個人未保存ならロール共通設定にフォールバック）。"""
    personal = get_mail_setting(_mgr_setting_key(user_id))
    return personal if not _is_empty_setting(personal) else get_mail_setting("管理職")


def _get_master_setting(user_id: int) -> dict:
    """マスタ用タブのメール設定を取得する（個人未保存ならロール共通設定にフォールバック）。"""
    personal = get_mail_setting(_master_setting_key(user_id))
    return personal if not _is_empty_setting(personal) else get_mail_setting("マスタ")


def _get_user_mail_setting(user_id: int) -> dict:
    """ユーザー個別のメール設定を取得する。

    Args:
        user_id: ユーザーID。

    Returns:
        dict: メール設定辞書。
    """
    return get_mail_setting(f"ユーザー_{user_id}")


def _get_user_body_template(user_id: int) -> str:
    """ユーザー個別の本文テンプレートを取得する。

    Args:
        user_id: ユーザーID。

    Returns:
        str: 本文テンプレート文字列。
    """
    setting = _get_user_mail_setting(user_id)
    body = setting.get("body_template", "").strip()
    return body if body else _USER_DEFAULT_BODY


@mail_report_bp.route("/user-preview")
def user_preview():
    """ユーザー用日報メールのプレビュー画面。

    Returns:
        str: プレビューHTML
    """
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))

    raw_date = request.args.get("date", "").strip()
    try:
        target_date = date.fromisoformat(raw_date)
        date_str = target_date.isoformat()
    except ValueError:
        target_date = date.today()
        date_str = target_date.isoformat()

    WEEKDAY_JA = ["月", "火", "水", "木", "金", "土", "日"]
    day_of_week = WEEKDAY_JA[target_date.weekday()]
    date_display = f"{target_date.year}/{target_date.month:02d}/{target_date.day:02d}"

    login_user = get_user_by_id(int(session["user_id"]))
    if not login_user:
        abort(404)

    uid = login_user["id"]
    setting = _get_user_mail_setting(uid)
    subject = _build_user_subject(login_user, target_date)
    body = _get_user_body_template(uid)
    mailto_url = _build_mailto(setting, subject)

    return render_template(
        "mail_user_preview.html",
        date_str=date_str,
        date_display=date_display,
        day_of_week=day_of_week,
        setting=setting,
        subject=subject,
        body=body,
        mailto_url=mailto_url,
        csrf_token=session.get("csrf_token", ""),
    )


@mail_report_bp.route("/save-user-address", methods=["POST"])
def save_user_address() -> object:
    """ユーザー用メールの宛先設定を保存する。

    Returns:
        object: プレビュー画面へのリダイレクト。
    """
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    uid = int(session["user_id"])
    current = _get_user_mail_setting(uid)
    save_mail_setting(
        role=f"ユーザー_{uid}",
        to_address=request.form.get("to_address", "").strip(),
        cc_address=request.form.get("cc_address", "").strip(),
        subject_template=current.get("subject_template", ""),
        body_template=current.get("body_template", ""),
        bcc_address=request.form.get("bcc_address", "").strip(),
    )

    date_str = request.form.get("date_str", "")
    return redirect(url_for("mail_report_bp.user_preview", date=date_str))


@mail_report_bp.route("/save-user-body", methods=["POST"])
def save_user_body() -> object:
    """ユーザー用メールの本文テンプレートを保存する。

    Returns:
        object: プレビュー画面へのリダイレクト。
    """
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))
    if request.form.get("csrf_token") != session.get("csrf_token"):
        abort(400)

    uid = int(session["user_id"])
    current = _get_user_mail_setting(uid)
    save_mail_setting(
        role=f"ユーザー_{uid}",
        to_address=current.get("to_address", ""),
        cc_address=current.get("cc_address", ""),
        subject_template=current.get("subject_template", ""),
        body_template=request.form.get("body_template", "").strip(),
        bcc_address=current.get("bcc_address", ""),
    )

    date_str = request.form.get("date_str", "")
    return redirect(url_for("mail_report_bp.user_preview", date=date_str))


@mail_report_bp.route("/download-user-eml", methods=["GET", "POST"])
def download_user_eml():
    """ユーザー用メールを.emlファイルとしてダウンロードする。

    POSTの場合はフォームから本文を受け取り、編集後の内容をEMLに反映する。
    GETの場合は保存済みテンプレートを用いる。

    Returns:
        Response: .emlファイルのダウンロードレスポンス。
    """
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))

    # POST時はCSRFを検証
    if request.method == "POST":
        if request.form.get("csrf_token") != session.get("csrf_token"):
            abort(400)

    raw_date = (request.values.get("date", "") or "").strip()
    try:
        target_date = date.fromisoformat(raw_date)
    except ValueError:
        target_date = date.today()

    login_user = get_user_by_id(int(session["user_id"]))
    if not login_user:
        abort(404)

    uid = login_user["id"]
    setting = _get_user_mail_setting(uid)
    subject = _build_user_subject(login_user, target_date)

    # POSTで本文が渡されている場合は、編集後の本文を使用
    posted_body = request.form.get("body", "").strip() if request.method == "POST" else ""
    body = posted_body if posted_body else _get_user_body_template(uid)

    eml_content = _build_eml(setting, subject, body)
    filename = f"daily_report_{target_date.isoformat()}_user.eml"

    return Response(
        eml_content,
        mimetype="message/rfc822",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )
