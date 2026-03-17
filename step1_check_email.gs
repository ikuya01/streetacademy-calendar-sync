/**
 * ストアカ予約通知メールの形式を確認するスクリプト
 *
 * 使い方:
 * 1. Google Apps Script (script.google.com) で新しいプロジェクトを作成
 * 2. このコードを貼り付けて checkStreetAcademyEmail() を実行
 * 3. ログ（表示 → ログ）でメールの件名・本文を確認
 */

function checkStreetAcademyEmail() {
  // ストアカからのメールを検索（直近のもの）
  var threads = GmailApp.search('from:no_reply_mail@street-academy.com subject:予約', 0, 5);

  if (threads.length === 0) {
    // 別パターンでも検索
    threads = GmailApp.search('from:street-academy.com', 0, 5);
  }

  if (threads.length === 0) {
    Logger.log('ストアカからのメールが見つかりませんでした。');
    Logger.log('検索条件を確認してください。');
    return;
  }

  Logger.log('=== 見つかったメール: ' + threads.length + '件 ===\n');

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    var msg = messages[messages.length - 1]; // スレッドの最新メッセージ

    Logger.log('--- メール ' + (i + 1) + ' ---');
    Logger.log('件名: ' + msg.getSubject());
    Logger.log('差出人: ' + msg.getFrom());
    Logger.log('日時: ' + msg.getDate());
    Logger.log('');
    Logger.log('【プレーンテキスト本文】');
    Logger.log(msg.getPlainBody());
    Logger.log('');
    Logger.log('========================\n');
  }
}
