"""管理職日報メール画面・設定ルート。"""
from __future__ import annotations

import urllib.parse
from datetime import date, timedelta

from flask import Blueprint, abort, redirect, render_template, request, session, url_for

from ..auth_helpers import is_privileged, is_master
from ..models import (
    get_accessible_users,
    get_daily_result,
    get_daily_comment,
    get_mail_setting,
    get_task_master,
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


def _next_workday(from_date: date) -> date:
    """翌稼働日（土日を除く）を返す。

    Args:
        from_date: 基準日

    Returns:
        date: 翌稼働日
    """
    d = from_date + timedelta(days=1)
    while d.weekday() >= 5:  # 5=土, 6=日
        d += timedelta(days=1)
    return d


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

    # 振り返り・要点
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

    # 翌稼働日の予定（同一作業名は1つにまとめる）
    next_day = _next_workday(target_date)
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
        "＜要点＞\n"
        f"{action}\n"
        "\n"
        "＜実施内容＞\n"
        f"{work_results}\n"
        "\n"
        "＜翌稼働日の達成目標＞\n"
        f"{next_schedule}\n"
        "\n"
        "以上になります。\n"
        "ご確認のほど、よろしくお願いいたします。"
    )

    return subject, body


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


def _build_master_body(
    login_user: dict, target_date: date, members: list[dict], greeting: str
) -> str:
    """マスタ用メール本文を動的生成する。

    大区分・中区分でグループ化した作業実績と、メンバー別AM/PMサマリ、
    振り返り、AI開発状況、次回予定を含む。

    Args:
        login_user: ログインユーザー情報
        target_date: 対象日
        members: アクセス可能なメンバー一覧
        greeting: 宛先挨拶文（設定から取得）

    Returns:
        str: メール本文
    """
    date_str = target_date.isoformat()
    dept = login_user.get("dept", "")
    login_id: int = login_user["id"]

    # 対象日の曜日・週開始日（週次スケジュール取得用）
    target_dow = target_date.weekday()
    week_start_str = (target_date - timedelta(days=target_dow)).isoformat()

    # 各メンバーの実績・タスクマスタ・週次スケジュールを収集
    member_data: list[dict] = []
    for member in members:
        uid = member["id"]
        result = get_daily_result(uid, date_str)
        comment_row = get_daily_comment(uid, date_str)
        task_master_list = get_task_master(uid)
        # task_name → {category_name, subcategory_name} マップ
        task_cat_map: dict[str, dict] = {
            t["task_name"]: {
                "category_name": t.get("category_name") or "",
                "subcategory_name": t.get("subcategory_name") or "",
            }
            for t in task_master_list
        }
        # 対象日の週次スケジュール（計画判定用）
        schedule_data = get_weekly_schedule(uid, week_start_str)
        scheduled_tasks: set[str] = {
            item["task_name"].strip()
            for slot in ("am", "pm")
            for item in schedule_data.get(target_dow, {}).get(slot, [])
            if item.get("task_name", "").strip()
        }
        member_data.append({
            "member": member,
            "result": result,
            "comment": comment_row,
            "task_cat_map": task_cat_map,
            "scheduled_tasks": scheduled_tasks,
        })

    # 全メンバーの標準時間合計（作業率・実績率計算用）
    total_std_hours = sum(
        float(md["member"].get("std_hours") or 8.0)
        for md in member_data
    )

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

    # 内訳%は総実績時間ベース、実績%は内訳の合計
    plan_rate = int(plan_hours / total_actual_hours * 100) if total_actual_hours > 0 else 0
    sudden_rate = int(sudden_hours / total_actual_hours * 100) if total_actual_hours > 0 else 0
    resc_rate = int(resc_hours / total_actual_hours * 100) if total_actual_hours > 0 else 0
    jisseki_rate = plan_rate + sudden_rate + resc_rate

    # 大区分→中区分→作業名でグループ化・時間集計
    cat_map: dict[str, dict[str, dict[str, float]]] = {}
    cat_order: list[str] = []
    subcat_order: dict[str, list[str]] = {}
    task_order: dict[str, dict[str, list[str]]] = {}

    for md in member_data:
        for slot in ("am", "pm"):
            for item in md["result"].get(slot, []):
                task = item.get("task_name", "").strip()
                hours = float(item.get("hours", 0.0))
                if not task or hours == 0.0:
                    continue
                cat_info = md["task_cat_map"].get(task, {})
                cat = cat_info.get("category_name") or "（区分なし）"
                subcat = cat_info.get("subcategory_name") or "（中区分なし）"
                if cat not in cat_map:
                    cat_map[cat] = {}
                    cat_order.append(cat)
                    subcat_order[cat] = []
                    task_order[cat] = {}
                if subcat not in cat_map[cat]:
                    cat_map[cat][subcat] = {}
                    subcat_order[cat].append(subcat)
                    task_order[cat][subcat] = []
                if task not in cat_map[cat][subcat]:
                    cat_map[cat][subcat][task] = 0.0
                    task_order[cat][subcat].append(task)
                cat_map[cat][subcat][task] += hours

    # 業務内容セクション
    content_lines: list[str] = ["業務内容 / 対応内容"]
    for cat in cat_order:
        content_lines.append(f"・{cat}")
        for subcat in subcat_order.get(cat, []):
            for task in task_order.get(cat, {}).get(subcat, []):
                t_hours = cat_map[cat][subcat][task]
                rate = int(t_hours / total_std_hours * 100) if total_std_hours > 0 else 0
                content_lines.append(f"  {subcat}　※{task}　{rate}%")

    # メンバー AM/PM サマリ
    member_lines: list[str] = []
    for md in member_data:
        name = md["member"]["name"]
        result = md["result"]
        am_tasks = list(dict.fromkeys(
            item["task_name"]
            for item in result.get("am", [])
            if item.get("task_name", "").strip() and float(item.get("hours", 0)) > 0
        ))
        pm_tasks = list(dict.fromkeys(
            item["task_name"]
            for item in result.get("pm", [])
            if item.get("task_name", "").strip() and float(item.get("hours", 0)) > 0
        ))
        am_str = "・".join(am_tasks) if am_tasks else "（なし）"
        pm_str = "・".join(pm_tasks) if pm_tasks else "（なし）"
        member_lines.append(f"{name}：{am_str} / {pm_str}")

    # マスタ自身の振り返り
    master_comment = get_daily_comment(login_id, date_str)
    reflection = master_comment.get("reflection", "").strip() or "（未入力）"

    # ＜開発状況＞: AI開発関連タスク（カテゴリ・中区分・作業名に「AI」を含む）
    ai_seen: set[str] = set()
    ai_lines: list[str] = []
    for md in member_data:
        for slot in ("am", "pm"):
            for item in md["result"].get(slot, []):
                task = item.get("task_name", "").strip()
                hours = float(item.get("hours", 0.0))
                if not task or hours == 0.0:
                    continue
                cat_info = md["task_cat_map"].get(task, {})
                cat = cat_info.get("category_name", "")
                subcat = cat_info.get("subcategory_name", "")
                if "AI" in cat or "AI" in subcat or "AI" in task:
                    line = f"  {md['member']['name']}：{task}　{hours}h"
                    if line not in ai_seen:
                        ai_seen.add(line)
                        ai_lines.append(line)
    ai_section = "\n".join(ai_lines) if ai_lines else "  （AI開発作業なし）"

    # ＜次回予定＞: マスタ自身の翌稼働日予定
    next_day = _next_workday(target_date)
    next_dow = next_day.weekday()
    next_week_start = (next_day - timedelta(days=next_dow)).isoformat()
    next_schedule_data = get_weekly_schedule(login_id, next_week_start)
    next_seen_list: list[str] = []
    for slot in ("am", "pm"):
        for item in next_schedule_data.get(next_dow, {}).get(slot, []):
            t = item.get("task_name", "").strip()
            if t and t not in next_seen_list:
                next_seen_list.append(t)
    next_schedule = "\n".join(f"・{t}" for t in next_seen_list) if next_seen_list else "（予定未入力）"

    # 本文組み立て
    parts: list[str] = []
    if greeting.strip():
        parts.append(greeting.strip())
    parts.extend([
        f"お疲れ様です。{dept}の業務報告となります。",
        "",
        "□予定：100%（作業計画：100%）",
        f"■実績：{jisseki_rate}%（計画：{plan_rate}%　突発：{sudden_rate}%　リスケ：{resc_rate}%　）",
        "",
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


def _build_mailto(setting: dict, subject: str, body: str) -> str:
    """mailto: URLを組み立てる。

    Args:
        setting: メール設定 dict（to_address, cc_address を含む）
        subject: メール件名
        body: メール本文

    Returns:
        str: mailto: スキームのURL文字列
    """
    params: dict[str, str] = {"subject": subject, "body": body}
    if setting.get("cc_address"):
        params["cc"] = setting["cc_address"]
    query = urllib.parse.urlencode(params, quote_via=urllib.parse.quote)
    to = urllib.parse.quote(setting.get("to_address", ""))
    return f"mailto:{to}?{query}"


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
    mgr_mailto = _build_mailto(mgr_setting, mgr_subject, mgr_body)

    # マスタ用: 動的生成（件名は曜日で自動判定、本文は大区分・中区分グループ化）
    dept = login_user.get("dept", "")
    members = get_accessible_users(login_user["id"], login_user["role"], dept)
    master_subject = _build_master_subject(dept, target_date)
    master_greeting = master_setting.get("body_template", "")
    master_body = _build_master_body(login_user, target_date, members, master_greeting)
    master_mailto = _build_mailto(master_setting, master_subject, master_body)

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
    )

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
