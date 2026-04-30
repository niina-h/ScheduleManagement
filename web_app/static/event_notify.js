/**
 * イベント開始10分前のWindows通知（ブラウザNotification API）
 *
 * 仕組み:
 *  1) ページロード時と5分ごとに /api/today-events を呼び、本日のイベント一覧を取得
 *  2) 各イベントについて「開始時刻 - 10分」の時点で発火する setTimeout を個別予約
 *  3) 時刻到来で new Notification(...) でWindows通知を発火
 *  4) 通知済みイベントは localStorage に記録して二重通知を防止
 *  5) イベント情報の再取得時は既存の予約を一旦クリアして再予約（編集・追加に追随）
 *
 * 外部通信なし。社内LAN（予定管理システム本体）との通信のみで完結する。
 */
(function () {
  'use strict';

  // 通知非対応ブラウザ（IE等）はスキップ
  if (!('Notification' in window)) return;

  // 未ログインページ（ログイン画面など）ではスキップ
  if (document.body && document.body.dataset.loggedIn !== '1') return;

  /** 「開始時刻のN分前」に通知する。今回は10分前。 */
  var NOTIFY_BEFORE_MIN = 10;
  var NOTIFY_BEFORE_MS = NOTIFY_BEFORE_MIN * 60 * 1000;

  /** 通知済みID集合（localStorage で当日分を保持し二重通知を防止）。 */
  var notifiedIds = new Set();
  var todayKey = 'eventNotified_' + new Date().toDateString();
  try {
    var stored = localStorage.getItem(todayKey);
    if (stored) notifiedIds = new Set(JSON.parse(stored));
    // 古い日付のキーを掃除
    for (var i = localStorage.length - 1; i >= 0; i--) {
      var k = localStorage.key(i);
      if (k && k.indexOf('eventNotified_') === 0 && k !== todayKey) {
        localStorage.removeItem(k);
      }
    }
  } catch (e) { /* localStorage 利用不可でも継続 */ }

  /**
   * 通知許可をユーザー操作後に1回だけリクエストする。
   * デフォルト状態の場合のみ呼ばれる。
   */
  (function requestPermissionOnFirstClick() {
    if (Notification.permission !== 'default') return;
    var handler = function () {
      Notification.requestPermission().catch(function () { /* noop */ });
      document.removeEventListener('click', handler);
    };
    document.addEventListener('click', handler, { once: true });
  })();

  /** イベント情報をAPIから取得する。失敗時は空配列。 */
  function fetchEvents() {
    return fetch('/api/today-events', { credentials: 'same-origin' })
      .then(function (res) { return res.ok ? res.json() : { events: [] }; })
      .then(function (data) { return (data && data.events) || []; })
      .catch(function () { return []; });
  }

  /**
   * 'YYYY-MM-DD' と 'HH:MM' (or 'HH:MM:SS') を結合して Date オブジェクトを返す。
   * 失敗時は null。
   */
  function parseEventStart(dateStr, timeStr) {
    if (!dateStr || !timeStr) return null;
    // ブラウザ互換のため "/" 区切りに変換し、ローカルタイムとして解釈させる
    var d = new Date(dateStr.replace(/-/g, '/') + ' ' + timeStr);
    if (isNaN(d.getTime())) return null;
    return d;
  }

  /** 通知済みID集合を localStorage に保存する。 */
  function persistNotifiedIds() {
    try {
      localStorage.setItem(todayKey, JSON.stringify(Array.from(notifiedIds)));
    } catch (e) { /* noop */ }
  }

  /** Windows通知を発火する。 */
  function fireNotification(ev) {
    if (notifiedIds.has(ev.id)) return;
    if (Notification.permission !== 'granted') return;
    try {
      new Notification('まもなく開始：' + ev.task_name, {
        body: '開始時刻 ' + ev.event_start_time + '（あと' + NOTIFY_BEFORE_MIN + '分）',
        tag: 'evt-' + ev.id,
        requireInteraction: false,
      });
    } catch (e) { /* 通知表示失敗時も継続 */ }
    notifiedIds.add(ev.id);
    persistNotifiedIds();
  }

  /** 予約済み setTimeout の管理（ev.id → timeoutId）。 */
  var scheduledTimeouts = new Map();

  /** 既存の予約をすべてクリアする（再取得時の再予約に備える）。 */
  function clearScheduled() {
    scheduledTimeouts.forEach(function (id) { clearTimeout(id); });
    scheduledTimeouts.clear();
  }

  /**
   * 1イベントの「開始10分前」発火 setTimeout を予約する。
   * 既に過去の場合・既通知済みの場合・時刻不正の場合はスキップ。
   */
  function scheduleNotification(ev) {
    if (notifiedIds.has(ev.id)) return;
    var start = parseEventStart(ev.start_date, ev.event_start_time);
    if (!start) return;
    var notifyAt = start.getTime() - NOTIFY_BEFORE_MS;
    var delay = notifyAt - Date.now();
    if (delay <= 0) return;  // 既に通知タイミングを過ぎているのでスキップ

    var timeoutId = setTimeout(function () {
      scheduledTimeouts.delete(ev.id);
      fireNotification(ev);
    }, delay);
    scheduledTimeouts.set(ev.id, timeoutId);
  }

  /** イベント一覧について全予約をリセットして再予約する。 */
  function rescheduleAll(events) {
    clearScheduled();
    events.forEach(scheduleNotification);
  }

  // 起動：イベント取得 + 予約
  fetchEvents().then(rescheduleAll);

  // 5分ごとにイベント情報を再取得（編集・追加への追随）し、予約も更新
  setInterval(function () {
    fetchEvents().then(rescheduleAll);
  }, 5 * 60 * 1000);
})();
