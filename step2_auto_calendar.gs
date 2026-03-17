/**
 * ストアカ予約・キャンセル通知メール → Googleカレンダー自動登録・アーカイブ
 *
 * 対応メール:
 *   A) クレカ決済済み予約
 *      From: no_reply_mail@street-academy.com
 *      件名: 【予約】「講座名」に予約が入りました
 *   B) 銀行振込の支払い完了
 *      From: information@street-academy.com
 *      件名: 「講座名」の予約が確定しました（お支払い完了）
 *   C) キャンセル通知（参加者0人→アーカイブ化）
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
  EVENT_COLOR: '7',           // ピーコック（青緑）
  ARCHIVED_EVENT_COLOR: '8',  // グラファイト（グレー）

  // 予約メール検索クエリ
  RESERVATION_QUERIES: [
    'from:no_reply_mail@street-academy.com subject:"予約が入りました" -subject:未払い',
    'from:information@street-academy.com subject:"予約が確定しました" subject:"お支払い完了"',
  ],

  // キャンセルメール検索クエリ
  CANCEL_QUERY: 'from:no_reply_mail@street-academy.com subject:"予約がキャンセルされました"',
};

// ========== 共通：日時パース ==========

/** 日時正規表現（年なし）- 半角・全角括弧両対応、全角チルダ対応 */
var DATE_REGEX = /開催日時[：:]\s*(\d{1,2})月(\d{1,2})日\s*[(（][日月火水木金土祝]+[)）]\s*(\d{1,2}):(\d{2})\s*[\u002D\u2013\u2014~〜\uff5e]\s*(\d{1,2}):(\d{2})/;

/** 日時正規表現（年あり） */
var DATE_REGEX_WITH_YEAR = /開催日時[：:]\s*(\d{4})年(\d{1,2})月(\d{1,2})日\s*[(（][日月火水木金土祝]+[)）]\s*(\d{1,2}):(\d{2})\s*[\u002D\u2013\u2014~〜\uff5e]\s*(\d{1,2}):(\d{2})/;

/**
 * メール本文から開催日時を抽出する
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

  var emailYear = emailDate.getFullYear();
  var month = parseInt(match[1]) - 1;
  var day = parseInt(match[2]);
  var startDate = new Date(emailYear, month, day, parseInt(match[3]), parseInt(match[4]));
  var endDate = new Date(emailYear, month, day, parseInt(match[5]), parseInt(match[6]));

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
    sendErrorAlert('processStreetAcademyEmails', 'カレンダーが見つかりません。CALENDAR_IDを確認してください。');
    return;
  }

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
          // アーカイブ済み（【開催なし】）のイベントがあれば復活させる
          restoreArchivedEvent(calendar, result);

          if (!isDuplicate(calendar, result)) {
            // ダブルブッキング検知
            checkTimeConflict(calendar, result);

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
        sendErrorAlert('processStreetAcademyEmails', '件名: ' + msg.getSubject() + '\nエラー: ' + e.message);
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
    // 個人情報を含まないよう件名のみログ出力
    Logger.log('日時パース失敗。件名: ' + subject);
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
 * 参加人数が0人 → イベントをアーカイブ化（【開催なし】+ グレー色）
 * 参加人数が1人以上 → 何もしない
 */
function processCancellationEmails() {
  var label = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
  var labelFilter = ' -label:' + CONFIG.PROCESSED_LABEL.replace(/\//g, '-');

  var calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) {
    sendErrorAlert('processCancellationEmails', 'カレンダーが見つかりません。');
    return;
  }

  var sinceDate = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000);
  var sinceStr = ' after:' + Utilities.formatDate(sinceDate, 'Asia/Tokyo', 'yyyy/MM/dd');

  var query = CONFIG.CANCEL_QUERY + labelFilter + sinceStr;
  var threads = GmailApp.search(query, 0, 20);

  if (threads.length === 0) return;

  Logger.log('キャンセルメール: ' + threads.length + '件');
  var archived = 0, kept = 0;

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    var msg = messages[messages.length - 1];

    try {
      var result = parseCancellationEmail(msg);

      if (result) {
        if (result.currentReservations === 0) {
          // 参加者0人 → アーカイブ化（削除ではなく色・タイトル変更）
          var done = archiveCalendarEvent(calendar, result.title, result.startDate, result.endDate);
          if (done) {
            Logger.log('✓ アーカイブ: ' + result.title + ' (' +
              Utilities.formatDate(result.startDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ') 参加者0人');
            archived++;
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
      sendErrorAlert('processCancellationEmails', '件名: ' + msg.getSubject() + '\nエラー: ' + e.message);
    }

    threads[i].addLabel(label);
  }

  if (archived > 0 || kept > 0) {
    Logger.log('キャンセル処理完了: アーカイブ' + archived + '件, 維持' + kept + '件');
  }
}

/**
 * キャンセルメールをパースする
 * 予約状況がパースできない場合は安全のためnullを返す（誤アーカイブ防止）
 */
function parseCancellationEmail(message) {
  var subject = message.getSubject();
  var body = message.getPlainBody();

  // --- 講座名の抽出 ---
  var titleMatch = body.match(/講座名[：:]\s*(.+)/);
  var title = null;
  if (titleMatch) {
    title = titleMatch[1]
      .replace(/\(https?:\/\/[^\)]+\)/, '')  // URL除去
      .replace(/\r/g, '')                     // \r除去
      .trim();
  }

  // フォールバック: 件名から
  if (!title) {
    var subjectMatch = subject.match(/「(.+?)」の予約がキャンセルされました/);
    if (subjectMatch) {
      title = subjectMatch[1]
        .replace(/[（(]\d{1,2}月\d{1,2}日.+$/, '')  // 半角・全角括弧の日時除去
        .trim();
    }
  }

  if (!title) title = subject;

  // --- 開催日時の抽出 ---
  var dateTime = parseDateTimeFromBody(body, message.getDate());
  if (!dateTime) return null;

  // --- 現在の予約人数の抽出 ---
  // パース失敗時はnullを返す（デフォルト0で誤アーカイブするのを防ぐ）
  var reservationMatch = body.match(/現在の予約状況[：:]\s*(\d+)\/(\d+)[席人]?[がの]?予約済み/);
  if (!reservationMatch) {
    Logger.log('⚠ 予約状況パース失敗 - 安全のためスキップ: ' + subject);
    return null;
  }

  return {
    title: '【ストアカ】' + title,
    startDate: dateTime.startDate,
    endDate: dateTime.endDate,
    currentReservations: parseInt(reservationMatch[1]),
  };
}


// ====================================================================
// カレンダー操作
// ====================================================================

/**
 * カレンダーイベントをアーカイブ化する（削除ではなくグレーアウト）
 * タイトルを「【開催なし】元タイトル」に変更し、色をグレーにする
 * @return {boolean}
 */
function archiveCalendarEvent(calendar, title, startDate, endDate) {
  var events = calendar.getEvents(startDate, endDate);
  for (var i = 0; i < events.length; i++) {
    if (events[i].getTitle() === title) {
      events[i].setTitle('【開催なし】' + title);
      events[i].setColor(CONFIG.ARCHIVED_EVENT_COLOR);
      events[i].setDescription('参加者0人のためアーカイブ\n' + (events[i].getDescription() || ''));
      return true;
    }
  }
  return false;
}

/**
 * アーカイブ済みイベントを復活させる（キャンセル後に再予約が入った場合）
 */
function restoreArchivedEvent(calendar, result) {
  var events = calendar.getEvents(result.startDate, result.endDate);
  for (var i = 0; i < events.length; i++) {
    if (events[i].getTitle() === '【開催なし】' + result.title) {
      events[i].setTitle(result.title);
      events[i].setColor(CONFIG.EVENT_COLOR);
      events[i].setDescription(result.description);
      Logger.log('✓ 復活: ' + result.title + ' (アーカイブから復元)');
      return;
    }
  }
}

/**
 * ダブルブッキング検知
 * 同じ時間帯に別のストアカ講座がないかチェックし、あればメールで警告
 */
function checkTimeConflict(calendar, result) {
  var events = calendar.getEvents(result.startDate, result.endDate);
  var conflicts = [];
  for (var i = 0; i < events.length; i++) {
    var evTitle = events[i].getTitle();
    // 自分自身・アーカイブ済みは除外
    if (evTitle !== result.title &&
        evTitle.indexOf('【ストアカ】') === 0 &&
        evTitle.indexOf('【開催なし】') === -1) {
      conflicts.push(evTitle);
    }
  }

  if (conflicts.length > 0) {
    var msg = '⚠ ダブルブッキング検知！\n\n' +
      '新規予約: ' + result.title + '\n' +
      '日時: ' + Utilities.formatDate(result.startDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') +
      ' - ' + Utilities.formatDate(result.endDate, 'Asia/Tokyo', 'HH:mm') + '\n\n' +
      '重複する既存講座:\n- ' + conflicts.join('\n- ');

    Logger.log(msg);
    GmailApp.sendEmail(
      Session.getEffectiveUser().getEmail(),
      '[ストアカ] ダブルブッキング警告',
      msg
    );
  }
}


// ========== エラー通知 ==========

/**
 * エラー発生時にメールで通知する
 */
function sendErrorAlert(functionName, errorMessage) {
  try {
    GmailApp.sendEmail(
      Session.getEffectiveUser().getEmail(),
      '[ストアカGAS] エラー発生: ' + functionName,
      '関数: ' + functionName + '\n' +
      '日時: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss') + '\n\n' +
      errorMessage
    );
  } catch (e) {
    Logger.log('エラー通知メール送信失敗: ' + e.message);
  }
}


// ========== ユーティリティ ==========

function isDuplicate(calendar, result) {
  var events = calendar.getEvents(result.startDate, result.endDate);
  for (var i = 0; i < events.length; i++) {
    var title = events[i].getTitle();
    if (title === result.title) return true;
    // アーカイブ済みイベントがあれば復活処理に任せるので重複とみなさない
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

  // 【開催なし】も削除
  var archivedEvents = calendar.getEvents(new Date(2026, 0, 1), new Date(2028, 0, 1), {search: '【開催なし】'});
  Logger.log(archivedEvents.length + '件の【開催なし】イベントを削除...');
  for (var i = 0; i < archivedEvents.length; i++) {
    archivedEvents[i].deleteEvent();
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
