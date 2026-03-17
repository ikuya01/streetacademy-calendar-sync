/**
 * ストアカ予約・キャンセル通知メール → Googleカレンダー自動登録・削除
 *
 * 対応メール:
 *   A) クレカ決済済み予約
 *      From: no_reply_mail@street-academy.com
 *      件名: 【予約】「講座名」に予約が入りました
 *   B) 銀行振込の支払い完了
 *      From: information@street-academy.com
 *      件名: 「講座名」の予約が確定しました（お支払い完了）
 *   C) キャンセル通知（参加者0人なら削除）
 *      From: no_reply_mail@street-academy.com
 *      件名: 【キャンセル】「講座名（日時）」の予約がキャンセルされました
 *
 *   ※ 未払い通知（件名に「未払い」を含む）はスキップ
 *
 * トリガー設定:
 *   関数: processStreetAcademyEmails / processCancellationEmails
 *   間隔: 15分おき推奨
 */

// ========== 設定 ==========
var CONFIG = {
  PROCESSED_LABEL: 'ストアカ/カレンダー登録済み',
  CALENDAR_ID: 'primary',
  // イベントの色: 7=ピーコック（青緑）
  EVENT_COLOR: '7',

  // 予約メール検索クエリ
  RESERVATION_QUERIES: [
    'from:no_reply_mail@street-academy.com subject:"予約が入りました" -subject:未払い',
    'from:information@street-academy.com subject:"予約が確定しました" subject:"お支払い完了"',
  ],

  // キャンセルメール検索クエリ
  CANCEL_QUERY: 'from:no_reply_mail@street-academy.com subject:"予約がキャンセルされました"',
};

// ========== 共通：日時パース ==========

/** 日時正規表現（年なし） */
var DATE_REGEX = /開催日時[：:]\s*(\d{1,2})月(\d{1,2})日\s*\([日月火水木金土祝]\)\s*(\d{1,2}):(\d{2})\s*[-\u002D\u2013\u2014~〜]\s*(\d{1,2}):(\d{2})/;

/** 日時正規表現（年あり） */
var DATE_REGEX_WITH_YEAR = /開催日時[：:]\s*(\d{4})年(\d{1,2})月(\d{1,2})日\s*\([日月火水木金土祝]\)\s*(\d{1,2}):(\d{2})\s*[-\u002D\u2013\u2014~〜]\s*(\d{1,2}):(\d{2})/;

/**
 * メール本文から開催日時を抽出する
 * @param {string} body プレーンテキスト本文
 * @param {Date} emailDate メール受信日時
 * @return {{startDate: Date, endDate: Date}|null}
 */
function parseDateTimeFromBody(body, emailDate) {
  var match = body.match(DATE_REGEX);

  if (!match) {
    var matchYear = body.match(DATE_REGEX_WITH_YEAR);
    if (matchYear) {
      var year = parseInt(matchYear[1]);
      var startDate = new Date(year, parseInt(matchYear[2]) - 1, parseInt(matchYear[3]),
        parseInt(matchYear[4]), parseInt(matchYear[5]));
      var endDate = new Date(year, parseInt(matchYear[2]) - 1, parseInt(matchYear[3]),
        parseInt(matchYear[6]), parseInt(matchYear[7]));
      if (endDate <= startDate) endDate.setDate(endDate.getDate() + 1);
      return { startDate: startDate, endDate: endDate };
    }
    return null;
  }

  // 年なし → メール受信年で推定
  var emailYear = emailDate.getFullYear();
  var month = parseInt(match[1]) - 1;
  var day = parseInt(match[2]);
  var startDate = new Date(emailYear, month, day, parseInt(match[3]), parseInt(match[4]));
  var endDate = new Date(emailYear, month, day, parseInt(match[5]), parseInt(match[6]));

  // メール受信日より2ヶ月以上前 → 翌年（12月→1月のケース）
  var twoMonthsBefore = new Date(emailDate.getTime() - 60 * 24 * 60 * 60 * 1000);
  if (startDate < twoMonthsBefore) {
    startDate.setFullYear(emailYear + 1);
    endDate.setFullYear(emailYear + 1);
  }

  if (endDate <= startDate) endDate.setDate(endDate.getDate() + 1);

  return { startDate: startDate, endDate: endDate };
}


// ====================================================================
// 予約メール処理
// ====================================================================

function processStreetAcademyEmails() {
  var label = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
  var labelFilter = ' -label:' + CONFIG.PROCESSED_LABEL.replace(/\//g, '-');

  var calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) {
    Logger.log('エラー: カレンダーが見つかりません。');
    return;
  }

  // 直近2日分を検索（after: はUTC基準のため余裕を持つ）
  var sinceDate = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000);
  var sinceStr = ' after:' + Utilities.formatDate(sinceDate, 'Asia/Tokyo', 'yyyy/MM/dd');

  var created = 0, skipped = 0, errors = 0;

  for (var q = 0; q < CONFIG.RESERVATION_QUERIES.length; q++) {
    var query = CONFIG.RESERVATION_QUERIES[q] + labelFilter + sinceStr;
    var threads = GmailApp.search(query, 0, 20);

    if (threads.length === 0) continue;
    Logger.log('予約クエリ' + (q + 1) + ': ' + threads.length + '件');

    for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      var msg = messages[messages.length - 1];

      try {
        var result = parseReservationEmail(msg);

        if (result) {
          if (!isDuplicate(calendar, result)) {
            var event = calendar.createEvent(result.title, result.startDate, result.endDate);
            if (CONFIG.EVENT_COLOR) event.setColor(CONFIG.EVENT_COLOR);
            event.setDescription(result.description);

            Logger.log('✓ 登録: ' + result.title + ' (' +
              Utilities.formatDate(result.startDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ')');
            created++;
          } else {
            skipped++;
          }
        } else {
          Logger.log('⚠ パース失敗: ' + msg.getSubject());
          errors++;
        }
      } catch (e) {
        Logger.log('✗ エラー: ' + msg.getSubject() + ' - ' + e.message);
        errors++;
      }

      // エラー時も必ずラベル付与（無限ループ防止）
      threads[i].addLabel(label);
    }
  }

  if (created > 0 || errors > 0) {
    Logger.log('予約処理完了: 登録' + created + '件, スキップ' + skipped + '件, エラー' + errors + '件');
  }
}

/**
 * 予約メールをパースする
 */
function parseReservationEmail(message) {
  var subject = message.getSubject();
  var body = message.getPlainBody();

  // 未払いメールが検索で漏れた場合のガード
  if (subject.indexOf('未払い') !== -1) return null;

  // --- 講座名の抽出（件名から） ---
  var title = null;
  var matchA = subject.match(/「(.+?)」に予約が入りました/);
  if (matchA) title = matchA[1].trim();

  if (!title) {
    var matchB = subject.match(/「(.+?)」の予約が確定しました/);
    if (matchB) title = matchB[1].trim();
  }

  if (!title) title = subject;

  // --- 開催日時の抽出 ---
  var dateTime = parseDateTimeFromBody(body, message.getDate());
  if (!dateTime) {
    Logger.log('日時パース失敗。本文先頭500文字:\n' + body.substring(0, 500));
    return null;
  }

  return {
    title: '【ストアカ】' + title,
    startDate: dateTime.startDate,
    endDate: dateTime.endDate,
    description: 'ストアカ予約通知メールから自動登録\n元メール件名: ' + subject,
  };
}


// ====================================================================
// キャンセルメール処理
// ====================================================================

/**
 * キャンセルメールを処理する
 * 参加人数が0人になった場合、対応するカレンダーイベントを削除する
 *
 * トリガー設定:
 *   関数: processCancellationEmails
 *   間隔: 15分おき
 */
function processCancellationEmails() {
  var label = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
  var labelFilter = ' -label:' + CONFIG.PROCESSED_LABEL.replace(/\//g, '-');

  var calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) {
    Logger.log('エラー: カレンダーが見つかりません。');
    return;
  }

  var sinceDate = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000);
  var sinceStr = ' after:' + Utilities.formatDate(sinceDate, 'Asia/Tokyo', 'yyyy/MM/dd');

  var query = CONFIG.CANCEL_QUERY + labelFilter + sinceStr;
  var threads = GmailApp.search(query, 0, 20);

  if (threads.length === 0) return;

  Logger.log('キャンセルメール: ' + threads.length + '件');
  var deleted = 0, kept = 0;

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    var msg = messages[messages.length - 1];

    try {
      var result = parseCancellationEmail(msg);

      if (result) {
        if (result.currentReservations === 0) {
          // 参加者0人 → カレンダーイベントを削除
          var removed = removeCalendarEvent(calendar, result.title, result.startDate, result.endDate);
          if (removed) {
            Logger.log('✓ 削除: ' + result.title + ' (' +
              Utilities.formatDate(result.startDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ') 参加者0人');
            deleted++;
          } else {
            Logger.log('- カレンダーにイベントなし: ' + result.title);
          }
        } else {
          Logger.log('- 維持: ' + result.title + ' (残り' + result.currentReservations + '人)');
          kept++;
        }
      } else {
        Logger.log('⚠ キャンセルメールのパース失敗: ' + msg.getSubject());
      }
    } catch (e) {
      Logger.log('✗ エラー: ' + msg.getSubject() + ' - ' + e.message);
    }

    threads[i].addLabel(label);
  }

  if (deleted > 0 || kept > 0) {
    Logger.log('キャンセル処理完了: 削除' + deleted + '件, 維持' + kept + '件');
  }
}

/**
 * キャンセルメールをパースする
 *
 * 件名: 【キャンセル】「講座名（3月21日(土) 9:00 - 10:30）」の予約がキャンセルされました
 * 本文:
 *   開催日時： 3月21日(土) 9:00 - 10:30
 *   現在の予約状況： 1/3席が予約済みです  ← この数字が0ならイベント削除
 */
function parseCancellationEmail(message) {
  var subject = message.getSubject();
  var body = message.getPlainBody();

  // --- 講座名の抽出（件名から） ---
  // 件名: 【キャンセル】「講座名（日時情報）」の予約がキャンセルされました
  // 講座名の後ろに日時が入っているので、本文の講座名ラベルから取得するほうが確実
  var titleMatch = body.match(/講座名[：:]\s*(.+)/);
  var title = null;
  if (titleMatch) {
    // URLやリンクテキストを除去
    title = titleMatch[1].replace(/\(https?:\/\/[^\)]+\)/, '').trim();
  }

  // フォールバック: 件名から（日時部分を含むが仕方ない）
  if (!title) {
    var subjectMatch = subject.match(/「(.+?)」の予約がキャンセルされました/);
    if (subjectMatch) {
      // 「講座名（日時）」から日時部分を除去
      title = subjectMatch[1].replace(/（\d{1,2}月\d{1,2}日.+$/, '').trim();
    }
  }

  if (!title) title = subject;

  // --- 開催日時の抽出 ---
  var dateTime = parseDateTimeFromBody(body, message.getDate());
  if (!dateTime) return null;

  // --- 現在の予約人数の抽出 ---
  // パターン: "現在の予約状況： 1/3席が予約済みです" or "現在の予約状況: 0/3席が予約済みです"
  var reservationMatch = body.match(/現在の予約状況[：:]\s*(\d+)\/(\d+)席が予約済み/);
  var currentReservations = 0;
  if (reservationMatch) {
    currentReservations = parseInt(reservationMatch[1]);
  }

  return {
    title: '【ストアカ】' + title,
    startDate: dateTime.startDate,
    endDate: dateTime.endDate,
    currentReservations: currentReservations,
  };
}

/**
 * カレンダーからイベントを削除する
 * @return {boolean} 削除できたかどうか
 */
function removeCalendarEvent(calendar, title, startDate, endDate) {
  var events = calendar.getEvents(startDate, endDate);
  for (var i = 0; i < events.length; i++) {
    if (events[i].getTitle() === title) {
      events[i].deleteEvent();
      return true;
    }
  }
  return false;
}


// ========== ユーティリティ ==========

function isDuplicate(calendar, result) {
  var events = calendar.getEvents(result.startDate, result.endDate);
  for (var i = 0; i < events.length; i++) {
    if (events[i].getTitle() === result.title) return true;
  }
  return false;
}

function getOrCreateLabel(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) label = GmailApp.createLabel(labelName);
  return label;
}


// ========== リセット用 ==========

/**
 * 誤登録されたカレンダーイベントを一括削除し、処理済みラベルを外す
 */
function resetAll() {
  var calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) {
    Logger.log('カレンダーが見つかりません。');
    return;
  }

  var events = calendar.getEvents(new Date(2026, 0, 1), new Date(2028, 0, 1), {search: '【ストアカ】'});
  Logger.log(events.length + '件の【ストアカ】イベントを削除...');
  for (var i = 0; i < events.length; i++) {
    Logger.log('  削除: ' + events[i].getTitle() + ' (' + events[i].getStartTime() + ')');
    events[i].deleteEvent();
  }

  var label = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL);
  if (label) {
    var threads = label.getThreads(0, 100);
    Logger.log(threads.length + '件のメールからラベルを外す...');
    for (var i = 0; i < threads.length; i++) {
      threads[i].removeLabel(label);
    }
  }

  Logger.log('リセット完了。');
}
