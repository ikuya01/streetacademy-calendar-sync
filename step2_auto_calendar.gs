/**
 * ストアカ予約通知メール → Googleカレンダー自動登録
 *
 * 対応する2パターンのメール:
 *   A) クレカ決済済み
 *      From: no_reply_mail@street-academy.com
 *      件名: 【予約】「講座名」に予約が入りました
 *   B) 銀行振込の支払い完了
 *      From: information@street-academy.com
 *      件名: 「講座名」の予約が確定しました（お支払い完了）
 *
 *   ※ 未払い通知（件名に「未払い」を含む）はスキップ
 *
 * 使い方:
 * 1. Google Apps Script (script.google.com) で新しいプロジェクトを作成
 * 2. このコードを貼り付け
 * 3. processStreetAcademyEmails() を手動実行してテスト
 * 4. 動作確認後、トリガーを設定（編集 → トリガー → 新しいトリガー）
 *    - 関数: processStreetAcademyEmails
 *    - イベントソース: 時間主導型
 *    - 時間ベースのトリガータイプ: 分ベースのタイマー
 *    - 間隔: 5分おき or 15分おき
 */

// ========== 設定 ==========
var CONFIG = {
  // 処理済みラベル名（処理済みメールに付けるGmailラベル）
  PROCESSED_LABEL: 'ストアカ/カレンダー登録済み',

  // カレンダーID（デフォルトカレンダーを使う場合は 'primary'）
  CALENDAR_ID: 'primary',

  // イベントの色（null で色指定なし）
  // 1:ラベンダー, 2:セージ, 3:ブドウ, 4:フラミンゴ, 5:バナナ,
  // 6:ミカン, 7:ピーコック, 8:グラファイト, 9:ブルーベリー, 10:バジル, 11:トマト
  EVENT_COLOR: '7', // ピーコック（青緑）

  // 検索クエリ（2パターン）
  SEARCH_QUERIES: [
    // A) クレカ決済済み（未払い除外）
    'from:no_reply_mail@street-academy.com subject:予約が入りました -subject:未払い',
    // B) 銀行振込の支払い完了
    'from:information@street-academy.com subject:予約が確定しました subject:お支払い完了',
  ],
};

// ========== メイン処理 ==========

function processStreetAcademyEmails() {
  var label = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
  var labelFilter = ' -label:' + CONFIG.PROCESSED_LABEL.replace(/\//g, '-');

  var calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) {
    Logger.log('エラー: カレンダーが見つかりません。CALENDAR_ID を確認してください。');
    return;
  }

  var created = 0;
  var skipped = 0;
  var errors = 0;

  // 直近1日分のメールだけ検索（トリガー5分おき想定、余裕を持って1日）
  var sinceDate = new Date(Date.now() - 24 * 60 * 60 * 1000);
  var sinceStr = ' after:' + Utilities.formatDate(sinceDate, 'Asia/Tokyo', 'yyyy/MM/dd');

  // 各検索クエリでメールを取得
  for (var q = 0; q < CONFIG.SEARCH_QUERIES.length; q++) {
    var query = CONFIG.SEARCH_QUERIES[q] + labelFilter + sinceStr;
    var threads = GmailApp.search(query, 0, 50);

    if (threads.length === 0) continue;

    Logger.log('クエリ' + (q + 1) + ': ' + threads.length + '件発見');

    for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      var msg = messages[messages.length - 1];

      try {
        var result = parseReservationEmail(msg);

        if (result) {
          if (!isDuplicate(calendar, result)) {
            var event = calendar.createEvent(
              result.title,
              result.startDate,
              result.endDate
            );

            if (CONFIG.EVENT_COLOR) {
              event.setColor(CONFIG.EVENT_COLOR);
            }

            if (result.description) {
              event.setDescription(result.description);
            }

            Logger.log('✓ 登録: ' + result.title + ' (' +
              Utilities.formatDate(result.startDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') +
              ' - ' +
              Utilities.formatDate(result.endDate, 'Asia/Tokyo', 'HH:mm') + ')');
            created++;
          } else {
            Logger.log('- スキップ（重複）: ' + result.title);
            skipped++;
          }
        } else {
          Logger.log('⚠ パース失敗: ' + msg.getSubject());
          errors++;
        }

        // 処理済みラベルを付与（パース失敗でも付ける＝無限ループ防止）
        threads[i].addLabel(label);

      } catch (e) {
        Logger.log('✗ エラー: ' + msg.getSubject() + ' - ' + e.message);
        errors++;
      }
    }
  }

  if (created === 0 && skipped === 0 && errors === 0) {
    Logger.log('新しい予約通知メールはありません。');
  } else {
    Logger.log('\n完了: 登録' + created + '件, スキップ' + skipped + '件, エラー' + errors + '件');
  }
}


// ========== メール解析 ==========

/**
 * 予約通知メールを解析して講座情報を抽出する
 *
 * 対応パターン:
 *   A) 件名: 【予約】「講座名」に予約が入りました（クレカ決済済み）
 *   B) 件名: 「講座名」の予約が確定しました（お支払い完了）（銀行振込完了）
 *
 * 本文の日時形式（共通）:
 *   開催日時： 3月22日(日) 9:00 - 10:30
 */
function parseReservationEmail(message) {
  var subject = message.getSubject();
  var body = message.getPlainBody();

  // --- 講座名の抽出（件名から） ---
  var title = null;

  // パターンA: 【予約】「講座名」に予約が入りました
  var matchA = subject.match(/「(.+?)」に予約が入りました/);
  if (matchA) {
    title = matchA[1].trim();
  }

  // パターンB: 「講座名」の予約が確定しました（お支払い完了）
  if (!title) {
    var matchB = subject.match(/「(.+?)」の予約が確定しました/);
    if (matchB) {
      title = matchB[1].trim();
    }
  }

  // フォールバック
  if (!title) {
    title = subject;
  }

  // --- 開催日時の抽出（本文から） ---
  // 実際のパターン: "開催日時： 3月22日(日) 9:00 - 10:30"
  var dateTimeMatch = body.match(
    /開催日時[：:]\s*(\d{1,2})月(\d{1,2})日\s*\([日月火水木金土祝]\)\s*(\d{1,2}):(\d{2})\s*[-\u2013\u2014~〜]\s*(\d{1,2}):(\d{2})/
  );

  if (!dateTimeMatch) {
    // 年付きフォールバック: "開催日時： 2026年3月22日(日) 9:00 - 10:30"
    var dateTimeMatchWithYear = body.match(
      /開催日時[：:]\s*(\d{4})年(\d{1,2})月(\d{1,2})日\s*\([日月火水木金土祝]\)\s*(\d{1,2}):(\d{2})\s*[-\u2013\u2014~〜]\s*(\d{1,2}):(\d{2})/
    );

    if (dateTimeMatchWithYear) {
      var year = parseInt(dateTimeMatchWithYear[1]);
      var month = parseInt(dateTimeMatchWithYear[2]) - 1;
      var day = parseInt(dateTimeMatchWithYear[3]);
      var startHour = parseInt(dateTimeMatchWithYear[4]);
      var startMin = parseInt(dateTimeMatchWithYear[5]);
      var endHour = parseInt(dateTimeMatchWithYear[6]);
      var endMin = parseInt(dateTimeMatchWithYear[7]);

      var startDate = new Date(year, month, day, startHour, startMin);
      var endDate = new Date(year, month, day, endHour, endMin);

      if (endDate <= startDate) {
        endDate.setDate(endDate.getDate() + 1);
      }

      return buildResult(title, startDate, endDate, subject);
    }

    Logger.log('日時パース失敗。本文先頭500文字:\n' + body.substring(0, 500));
    return null;
  }

  // 年なしパターンの処理
  // メールの受信日から年を推定する（現在時刻ではなくメール受信時点で判定）
  var emailDate = message.getDate();
  var emailYear = emailDate.getFullYear();
  var month = parseInt(dateTimeMatch[1]) - 1; // 0-indexed
  var day = parseInt(dateTimeMatch[2]);
  var startHour = parseInt(dateTimeMatch[3]);
  var startMin = parseInt(dateTimeMatch[4]);
  var endHour = parseInt(dateTimeMatch[5]);
  var endMin = parseInt(dateTimeMatch[6]);

  // まずメール受信年で日付を作る
  var startDate = new Date(emailYear, month, day, startHour, startMin);
  var endDate = new Date(emailYear, month, day, endHour, endMin);

  // メール受信日より2ヶ月以上前の日付なら翌年と判定
  // （例: 12月に届いた1月の講座 → 翌年1月）
  var twoMonthsBefore = new Date(emailDate.getTime() - 60 * 24 * 60 * 60 * 1000);
  if (startDate < twoMonthsBefore) {
    startDate.setFullYear(emailYear + 1);
    endDate.setFullYear(emailYear + 1);
  }

  // 終了が開始より前なら翌日（深夜またぎ）
  if (endDate <= startDate) {
    endDate.setDate(endDate.getDate() + 1);
  }

  return buildResult(title, startDate, endDate, subject);
}

/**
 * 結果オブジェクトを組み立てる
 */
function buildResult(title, startDate, endDate, subject) {
  return {
    title: '【ストアカ】' + title,
    startDate: startDate,
    endDate: endDate,
    description: 'ストアカ予約通知メールから自動登録\n元メール件名: ' + subject,
  };
}


// ========== ユーティリティ ==========

/**
 * カレンダーに同じタイトル・時間のイベントがないかチェック
 */
function isDuplicate(calendar, result) {
  var events = calendar.getEvents(result.startDate, result.endDate);
  for (var i = 0; i < events.length; i++) {
    if (events[i].getTitle() === result.title) {
      return true;
    }
  }
  return false;
}

/**
 * Gmailラベルを取得（なければ作成）
 */
function getOrCreateLabel(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}


// ========== リセット用 ==========

/**
 * 誤登録されたカレンダーイベントを一括削除し、処理済みラベルを外す
 * ※ 初回設定ミスの修正用。通常は使わない
 */
function resetAll() {
  // 1. 【ストアカ】で始まるカレンダーイベントを全削除
  var calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  var startRange = new Date(2026, 0, 1); // 2026年1月1日
  var endRange = new Date(2028, 0, 1);   // 2028年1月1日
  var events = calendar.getEvents(startRange, endRange, {search: '【ストアカ】'});

  Logger.log(events.length + '件の【ストアカ】イベントを削除します...');
  for (var i = 0; i < events.length; i++) {
    Logger.log('  削除: ' + events[i].getTitle() + ' (' + events[i].getStartTime() + ')');
    events[i].deleteEvent();
  }

  // 2. 処理済みラベルを外す
  var label = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL);
  if (label) {
    var threads = label.getThreads();
    Logger.log(threads.length + '件のメールからラベルを外します...');
    for (var i = 0; i < threads.length; i++) {
      threads[i].removeLabel(label);
    }
  }

  Logger.log('リセット完了。processStreetAcademyEmails を再実行してください。');
}


// ========== テスト用 ==========

/**
 * 両パターンのパーステスト
 */
function testParse() {
  // パターンA: クレカ決済済み
  var testSubjectA = '【予約】「AI×副業✨初心者・在宅OK✨業務自動化で月+10万円最速実現」に予約が入りました';
  var testBodyA = [
    '予約者情報:',
    '講座名： AI×副業(https://www.street-academy.com/myclass/173483)',
    '開催日時： 3月15日(日) 9:00 - 10:30',
    '予約した生徒： 片山 健',
  ].join('\n');

  // パターンB: 銀行振込完了
  var testSubjectB = '「AI×副業✨初心者・在宅OK✨業務自動化で月+10万円最速実現」の予約が確定しました（お支払い完了）';
  var testBodyB = [
    '予約者情報:',
    '講座名： AI×副業',
    '開催日時： 3月21日(土) 9:00 - 10:30',
    '予約した生徒： 大内 百合',
  ].join('\n');

  Logger.log('=== パターンA: クレカ決済 ===');
  testParseOne(testSubjectA, testBodyA);

  Logger.log('\n=== パターンB: 銀行振込完了 ===');
  testParseOne(testSubjectB, testBodyB);
}

function testParseOne(subject, body) {
  // 件名パース
  var matchA = subject.match(/「(.+?)」に予約が入りました/);
  var matchB = subject.match(/「(.+?)」の予約が確定しました/);
  var title = matchA ? matchA[1] : (matchB ? matchB[1] : 'パース失敗');
  Logger.log('講座名: ' + title);

  // 日時パース
  var dateTimeMatch = body.match(
    /開催日時[：:]\s*(\d{1,2})月(\d{1,2})日\s*\([日月火水木金土祝]\)\s*(\d{1,2}):(\d{2})\s*[-\u2013\u2014~〜]\s*(\d{1,2}):(\d{2})/
  );

  if (dateTimeMatch) {
    var now = new Date();
    Logger.log('日時パース成功: ' +
      dateTimeMatch[1] + '月' + dateTimeMatch[2] + '日 ' +
      dateTimeMatch[3] + ':' + dateTimeMatch[4] + ' - ' +
      dateTimeMatch[5] + ':' + dateTimeMatch[6] +
      ' (年: ' + now.getFullYear() + ')');
  } else {
    Logger.log('日時パース失敗');
  }
}
