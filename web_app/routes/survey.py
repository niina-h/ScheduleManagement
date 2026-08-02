"""web_app/routes/survey.py - ログイン時アンケート（初回のみ回答）"""
from __future__ import annotations

import secrets

from flask import Blueprint, abort, flash, redirect, render_template, request, session, url_for

from ..auth_helpers import is_master
from ..models import get_survey_answers, has_answered_survey, save_survey_answer

survey_bp = Blueprint("survey_bp", __name__)

# 「他部署・他社への推奨」の選択肢（フォーム側のvalueと一致させる）
_RECOMMEND_CHOICES: frozenset[str] = frozenset({"はい", "いいえ", "どちらとも言えない"})


def _csrf_ok() -> bool:
    """POSTフォームのCSRFトークンがセッションのものと一致するか検証する。

    Returns:
        bool: 一致すれば True。
    """
    token = request.form.get("csrf_token", "")
    return bool(token) and secrets.compare_digest(token, session.get("csrf_token", ""))


@survey_bp.route("/survey", methods=["GET", "POST"])
def survey() -> str:
    """ログイン時アンケートの表示・回答受付を行う。

    GET  : 未回答のユーザーにアンケートフォームを表示する。回答済みなら週間予定へ。
    POST : 回答内容を検証して保存し、週間予定へリダイレクトする。

    Returns:
        str: GETはHTMLレスポンス、POSTはリダイレクトレスポンス。
    """
    user_id = session.get("user_id")
    if not user_id:
        return redirect(url_for("auth.login"))

    if has_answered_survey(int(user_id)):
        return redirect(url_for("schedule.weekly"))

    if request.method == "POST":
        if not _csrf_ok():
            flash("セッションが無効です。もう一度お試しください", "danger")
            return redirect(url_for("survey_bp.survey"))

        usability_raw = request.form.get("usability", "")
        merit = request.form.get("merit", "")
        demerit = request.form.get("demerit", "")
        requested_feature = request.form.get("requested_feature", "")
        recommend = request.form.get("recommend", "")
        recommended_dept = request.form.get("recommended_dept", "")

        if usability_raw not in {"1", "2", "3", "4", "5"}:
            flash("使いやすさは1〜5から選択してください", "warning")
            return redirect(url_for("survey_bp.survey"))
        if recommend not in _RECOMMEND_CHOICES:
            flash("他部署・他社への推奨についてお答えください", "warning")
            return redirect(url_for("survey_bp.survey"))

        save_survey_answer(
            user_id=int(user_id),
            usability=int(usability_raw),
            merit=merit,
            demerit=demerit,
            requested_feature=requested_feature,
            recommend=recommend,
            recommended_dept=recommended_dept,
        )
        flash("アンケートにご協力いただきありがとうございました", "success")
        return redirect(url_for("schedule.weekly"))

    return render_template("survey.html")


@survey_bp.route("/survey/skip", methods=["POST"])
def survey_skip() -> str:
    """アンケートを「次回回答する」で見送り、今回のログインセッション中は再表示しない。

    回答自体は保存しないため、ログアウト・再ログイン時には再度表示される。

    Returns:
        str: 週間予定へのリダイレクト。
    """
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))
    if not _csrf_ok():
        flash("セッションが無効です。もう一度お試しください", "danger")
        return redirect(url_for("survey_bp.survey"))

    session["survey_skipped"] = True
    return redirect(url_for("schedule.weekly"))


@survey_bp.route("/survey/results")
def survey_results() -> str:
    """アンケート回答の一覧を表示する（システム管理者のみ）。

    誰がいつ回答したか（氏名・所属・回答日時）を含めて新しい順に表示する。

    Returns:
        str: レンダリングされたHTMLレスポンス。
    """
    if not session.get("user_id"):
        return redirect(url_for("auth.login"))
    if not is_master(session.get("user_role", "")):
        abort(403)

    answers = get_survey_answers()
    return render_template("survey_results.html", answers=answers)
