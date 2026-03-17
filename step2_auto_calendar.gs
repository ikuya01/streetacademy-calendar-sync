/**
 * ストアカ予約通知メール → Googleカレンダー自動登録
 *
 * 実際のメール形式に基づいて作成:
 *   件名: 【予約】「講座名」に予約が入りました（未払い）
 *   本文: 開催日時： 3月22日(日) 9:00 - 10:30
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

  // 検索クエリ（予約通知メールを絞り込む）
  SEARCH_QUERY: 'from:no_reply_mail@street-academy.com subject:予約が入りました',
};

// ========== メイン処理 ==========

function processStreetAcademyEmails() {
  var label = getOrCreateLabel(CONFIG.PROCESSED_LABEL);

  // 未処理のストアカ予約メールを検索
  // Gmailラベルのスラッシュはそのまま使える
  var query = CONFIG.SEARCH_QUERY + ' -label:' + CONFIG.PROCESSED_LABEL.replace(/\//g, '-');
  var threads = GmailApp.search(query, 0, 50);

  if (threads.length === 0) {
    Logger.log('新しい予約通知メールはありません。');
    return;
  }

  Logger.log(threads.length + '件の未処理メールを発見。');

  var calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) {
    Logger.log('エラー: カレンダーが見つかりません。CALENDAR_ID を確認してください。');
    return;
  }

  var created = 0;
  var skipped = 0;
  var errors = 0;

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

  Logger.log('\n完了: 登録' + created + '件, スキップ' + skipped + '件, エラー' + errors + '件');
}


// ========== メール解析 ==========

/**
 * 予約通知メールを解析して講座情報を抽出する
 *
 * 実際のメール形式:
 *   件名: 【予約】「AI×副業✨初心者・在宅OK✨業務自動化で月+10万円最速実現」に予約が入りました（未払い）
 *   本文（プレーンテキスト部）:
 *     講座名： class_detail.classname(URL)  ← リンクテキストは件名から取得するほうが確実
 *     開催日時： 3月22日(日) 9:00 - 10:30
 *     場所:　オンライン
 */
function parseReservationEmail(message) {
  var subject = message.getSubject();
  var body = message.getPlainBody();

  // --- 講座名の抽出（件名から） ---
  // 件名パターン: 【予約】「講座名」に予約が入りました（未払い）
  var titleMatch = subject.match(/「(.+?)」に予約が入りました/);
  if (!titleMatch) {
    // 「」なしのパターンにもフォールバック
    titleMatch = subject.match(/【予約】(.+?)に予約が入りました/);
  }
  var title = titleMatch ? titleMatch[1].trim() : subject;

  // --- 開催日時の抽出（本文から） ---
  // 実際のパターン: "開催日時： 3月22日(日) 9:00 - 10:30"
  //              or: "開催日時: 3月22日(日) 9:00 - 10:30"
  var dateTimeMatch = body.match(
    /開催日時[：:]\s*(\d{1,2})月(\d{1,2})日\s*\([日月火水木金土]\)\s*(\d{1,2}):(\d{2})\s*[-\u2013\u2014~〜]\s*(\d{1,2}):(\d{2})/
  );

  if (!dateTimeMatch) {
    // 年付きパターンにもフォールバック: "開催日時： 2026年3月22日(日) 9:00 - 10:30"
    var dateTimeMatchWithYear = body.match(
      /開催日時[：:]\s*(\d{4})年(\d{1,2})月(\d{1,2})日\s*\([日月火水木金土]\)\s*(\d{1,2}):(\d{2})\s*[-\u2013\u2014~〜]\s*(\d{1,2}):(\d{2})/
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
  var now = new Date();
  var year = now.getFullYear();
  var month = parseInt(dateTimeMatch[1]) - 1; // 0-indexed
  var day = parseInt(dateTimeMatch[2]);
  var startHour = parseInt(dateTimeMatch[3]);
  var startMin = parseInt(dateTimeMatch[4]);
  var endHour = parseInt(dateTimeMatch[5]);
  var endMin = parseInt(dateTimeMatch[6]);

  var startDate = new Date(year, month, day, startHour, startMin);
  var endDate = new Date(year, month, day, endHour, endMin);

  // 過去の日付なら来年と判定（例: 12月に届いた1月の講座）
  var oneWeekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  if (startDate < oneWeekAgo) {
    startDate.setFullYear(year + 1);
    endDate.setFullYear(year + 1);
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


// ========== テスト用 ==========

/**
 * 実際のメール形式に合わせたパーステスト
 */
function testParse() {
  var testBody = [
    'ikuya 先生さん',
    '',
    'ストアカ運営事務局からのお知らせです。',
    '',
    'ikuya 先生さんの以下の講座に銀行振込（未払い）で予約が入りました。',
    '',
    '予約者情報:',
    '講座名： AI×副業✨初心者・在宅OK✨業務自動化で月+10万円最速実現(https://www.street-academy.com/myclass/173483)',
    '開催日時： 3月22日(日) 9:00 - 10:30',
    '予約した生徒： Aki Sano',
    '申込み人数：1人',
    '場所:　オンライン',
  ].join('\n');

  var testSubject = '【予約】「AI×副業✨初心者・在宅OK✨業務自動化で月+10万円最速実現」に予約が入りました（未払い）';

  // 件名パース
  var titleMatch = testSubject.match(/「(.+?)」に予約が入りました/);
  Logger.log('講座名: ' + (titleMatch ? titleMatch[1] : 'パース失敗'));

  // 日時パース
  var dateTimeMatch = testBody.match(
    /開催日時[：:]\s*(\d{1,2})月(\d{1,2})日\s*\([日月火水木金土]\)\s*(\d{1,2}):(\d{2})\s*[-\u2013\u2014~〜]\s*(\d{1,2}):(\d{2})/
  );

  if (dateTimeMatch) {
    var now = new Date();
    Logger.log('パース成功:');
    Logger.log('  月: ' + dateTimeMatch[1]);
    Logger.log('  日: ' + dateTimeMatch[2]);
    Logger.log('  開始: ' + dateTimeMatch[3] + ':' + dateTimeMatch[4]);
    Logger.log('  終了: ' + dateTimeMatch[5] + ':' + dateTimeMatch[6]);
    Logger.log('  年（推定）: ' + now.getFullYear());

    var startDate = new Date(now.getFullYear(), parseInt(dateTimeMatch[1]) - 1, parseInt(dateTimeMatch[2]),
      parseInt(dateTimeMatch[3]), parseInt(dateTimeMatch[4]));
    var endDate = new Date(now.getFullYear(), parseInt(dateTimeMatch[1]) - 1, parseInt(dateTimeMatch[2]),
      parseInt(dateTimeMatch[5]), parseInt(dateTimeMatch[6]));

    Logger.log('  開始日時: ' + startDate);
    Logger.log('  終了日時: ' + endDate);
  } else {
    Logger.log('日時パース失敗');
    Logger.log('本文: ' + testBody);
  }
}
