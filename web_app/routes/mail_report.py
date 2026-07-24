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
    get_daily_result,
    get_daily_comment,
    get_events_for_user_date,
    get_global_task_category_map,
    get_mail_setting,
    get_next_business_day,
    get_task_master,
    get_weekly_leave,
    get_weekly_schedule,
    save_mail_setting,
    get_user_by_id,
)

mail_report_bp = Blueprint("mail_report_bp", __name__, url_prefix="/mail-report")


def _require_privileged() -> None | object:
    """管理職またはマスタでなければリダイレクト／403を返す。

    Returns:
        None: 権限チェック通過時
        redirect / abort: 権限不足時
    """
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))
    if not is_privileged(session.get("user_role", "")):
        abort(403)
    return None


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
    # 実施内容（本日の作業実績 - 同一作業名は時間を合算）
    result = get_daily_result(uid, date_str)
    work_totals: dict[str, float] = {}
    work_order: list[str] = []
    for slot in ("am", "pm"):
        for item in result.get(slot, []):
            task = item.get("task_name", "").strip()
            hours = float(item.get("hours", 0.0))
            if task:
                if task not in work_totals:
                    work_order.append(task)
                work_totals[task] = work_totals.get(task, 0.0) + hours
    work_lines: list[str] = [f"・{t}　{work_totals[t]}h" for t in work_order]
    work_results = "\n".join(work_lines) if work_lines else "（実績なし）"

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
    next_schedule = "\n".join(f"・{t}" for t in next_seen) if next_seen else "（予定未入力）"

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
        scheduled_tasks: set[str] = {
            item["task_name"].strip()
            for slot in ("am", "pm")
            for item in schedule_data.get(target_dow, {}).get(slot, [])
            if item.get("task_name", "").strip()
        }
        # 当日の休暇種別（1日有休 / AM半休 / PM半休 / 特休 / 祝日 / その他休み）
        leave_type: str = get_weekly_leave(uid, week_start_str).get(target_dow, "")
        member_data.append({
            "member": member,
            "result": result,
            "comment": comment_row,
            "task_cat_map": task_cat_map,
            "scheduled_tasks": scheduled_tasks,
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
    # 当日以前に完了したタスクは過去の業務報告で既に表示済みなので本文から除外する。
    member_ids = [m["id"] for m in members]
    target_date_str: str = target_date.isoformat()
    project_tasks = [
        t for t in get_all_project_tasks(user_ids=member_ids)
        if not t.get("is_event", 0)
        and not (
            t.get("status") == "完了"
            and (t.get("end_date") or "") <= target_date_str
        )
    ]
    # タスクを重複除去しつつ (大区分, 中区分, 状態) 付きで平坦化する
    # （画面の絞り込みと同じスコープ＝担当メンバーのタスクのみ、既に project_tasks で確定済み）。
    seen_task_names: set[str] = set()
    flat_tasks: list[dict] = []
    for pt in project_tasks:
        tname = pt["task_name"]
        if tname in seen_task_names:
            continue
        seen_task_names.add(tname)
        flat_tasks.append({
            "name": tname,
            "status": pt.get("status", ""),
            "cat_name": pt.get("category_name") or "",
            "sub_name": pt.get("subcategory_name") or "",
        })

    # 状態別に件数を集計し、遅れが目立つ順（遅れ→着手→未着手→順調→完了→停止）に並べる。
    # 「どのくらい作業があって、何が遅れているか」を先頭のサマリー行で即座に把握できるようにする。
    _STATUS_ORDER = ["遅れ", "着手", "未着手", "順調", "完了", "停止"]
    tasks_by_status: dict[str, list[dict]] = {s: [] for s in _STATUS_ORDER}
    for t in flat_tasks:
        tasks_by_status.setdefault(t["status"] or "未着手", []).append(t)

    status_counts = {s: len(tasks_by_status.get(s, [])) for s in _STATUS_ORDER}
    total_count = len(flat_tasks)
    delay_count = status_counts.get("遅れ", 0)

    def _format_task_line(t: dict) -> str:
        """1タスクを「タスク名（大区分／中区分）」形式の1行に整形する。

        区分が未設定（大区分・中区分ともに空）の場合は括弧書きを省略する。
        """
        cat_label = t["cat_name"]
        if t["sub_name"]:
            cat_label = f"{cat_label}／{t['sub_name']}" if cat_label else t["sub_name"]
        if not cat_label:
            return f"　・{t['name']}"
        return f"　・{t['name']}（{cat_label}）"

    content_lines: list[str] = []
    content_lines.append(
        f"全{total_count}件（遅れ{status_counts.get('遅れ', 0)}／着手{status_counts.get('着手', 0)}／"
        f"未着手{status_counts.get('未着手', 0)}／順調{status_counts.get('順調', 0)}／"
        f"完了{status_counts.get('完了', 0)}／停止{status_counts.get('停止', 0)}）"
    )
    for status in _STATUS_ORDER:
        items = tasks_by_status.get(status, [])
        if not items:
            continue
        content_lines.append(f"【{status}】{len(items)}件")
        for t in items:
            content_lines.append(_format_task_line(t))

    # メンバー AM/PM サマリ
    # 除外条件: 定例作業（大区分「定例」or 中区分「定例作業」）、AM1行目(idx=0)、PM最終行(idx=4)
    def _is_routine(task_name: str, cat_map: dict) -> bool:
        """定例作業かどうか判定する。"""
        info = cat_map.get(task_name.strip(), {})
        return (info.get("category_name") in ("定例", "定例作業")
                or info.get("subcategory_name") in ("定例", "定例作業"))

    # 終日休暇（出勤なし）と判定する種別
    FULL_LEAVE_TYPES: set[str] = {"1日有休", "特休", "その他休み", "祝日"}

    member_lines: list[str] = []
    for md in member_data:
        name = md["member"]["name"]
        result = md["result"]
        tcm = md["task_cat_map"]
        leave = md.get("leave_type", "")

        # 終日休暇の場合：休暇種別のみを表示
        if leave in FULL_LEAVE_TYPES:
            member_lines.append(f"{name}：{leave}")
            continue

        am_tasks = list(dict.fromkeys(
            item["task_name"]
            for idx, item in enumerate(result.get("am", []))
            if item.get("task_name", "").strip() and float(item.get("hours", 0)) > 0
            and idx != 0  # AM1行目を除外
            and not _is_routine(item["task_name"], tcm)
        ))
        pm_tasks = list(dict.fromkeys(
            item["task_name"]
            for idx, item in enumerate(result.get("pm", []))
            if item.get("task_name", "").strip() and float(item.get("hours", 0)) > 0
            and idx != 4  # PM最終行(5枠目)を除外
            and not _is_routine(item["task_name"], tcm)
        ))
        am_str = "/ ".join(am_tasks) if am_tasks else "（なし）"
        pm_str = "/ ".join(pm_tasks) if pm_tasks else "（なし）"

        # 半休の場合：休んでいる側を休暇種別表記、もう片方は通常通り
        if leave == "AM半休":
            member_lines.append(f"{name}：AM半休 / {pm_str}")
        elif leave == "PM半休":
            member_lines.append(f"{name}：{am_str} / PM半休")
        elif am_str == pm_str:
            member_lines.append(f"{name}：{am_str}")
        else:
            member_lines.append(f"{name}：{am_str} / {pm_str}")

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
    #   形式: 「  中区分　タスク名　（氏名：進捗X%、状態）」
    user_name_by_id: dict[int, str] = {m["id"]: m["name"] for m in members}

    def _dev_assignee_names(pt: dict) -> str:
        """タスクの担当者名（親ならメンバー、子なら担当者1・2）を「、」区切りで返す。"""
        member_ids_raw = (pt.get("member_ids") or "").strip()
        if member_ids_raw:
            names = [
                user_name_by_id[int(s)] for s in member_ids_raw.split(",")
                if s.strip().isdigit() and int(s) in user_name_by_id
            ]
            return "、".join(names) if names else "担当者未設定"
        names = []
        if pt.get("assigned_name"):
            names.append(pt.get("assigned_last_name") or pt.get("assigned_name"))
        if pt.get("assigned_name_2"):
            names.append(pt.get("assigned_last_name_2") or pt.get("assigned_name_2"))
        return "、".join(names) if names else "担当者未設定"

    dev_tasks = [pt for pt in project_tasks if (pt.get("category_name") or "") == "開発"]
    ai_lines: list[str] = []
    for pt in dev_tasks:
        label = (pt.get("subcategory_name") or "").strip() or "開発"
        assignees = _dev_assignee_names(pt)
        status = pt.get("status", "") or "未着手"
        progress = pt.get("progress", 0) or 0
        ai_lines.append(f"  {label}　{pt['task_name']}　（{assignees}：進捗{progress}%、{status}）")
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
            if t and t not in next_seen_list:
                if master_tcm.get(t, {}).get("subcategory_name") != "定例作業":
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
        "□予定：100%（作業計画：100%）",
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
        dept = login_user.get("dept", "")
        members = get_accessible_users(login_user["id"], login_user["role"], dept)
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
    mgr_setting = get_mail_setting("管理職")
    master_setting = get_mail_setting("マスタ")

    # 管理職用: 固定テンプレート
    mgr_subject, mgr_body = _build_mgr_self_body(login_user, target_date)
    mgr_mailto = _build_mailto(mgr_setting, mgr_subject)

    # マスタ用: 動的生成（件名は曜日で自動判定、本文は大区分・中区分グループ化）
    dept = login_user.get("dept", "")
    members = get_accessible_users(login_user["id"], login_user["role"], dept)
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

    dept = login_user.get("dept", "")
    members = get_accessible_users(login_user["id"], login_user["role"], dept)
    master_subject = _build_master_subject(_resolve_master_dept(login_user, members), target_date)
    master_greeting = get_mail_setting("マスタ").get("body_template", "")
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

    current = get_mail_setting(role)
    save_mail_setting(
        role=role,
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
