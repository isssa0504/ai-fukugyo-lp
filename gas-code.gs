// ===== Google Apps Script =====
// 使い方:
// 1. Google Drive → 新規 → Google Apps Script
// 2. このコードを貼り付け
// 3. SHEET_ID を自分のスプレッドシートIDに変更
// 4. 「デプロイ」→「新しいデプロイ」→ 種類:ウェブアプリ → アクセス:全員 → デプロイ
// 5. 表示されたURLをLP内の GAS_URL に貼り付け

const SHEET_ID = '1srAhDKrPS75qZxcW2GElnXelATYN1KqG3kAa2KHdpDo';
const SHEET_NAME = '予約一覧';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);

    // シートがなければ作成してヘッダーを入れる
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['受付日時', 'お名前', 'LINE名', '候補①', '候補②', '候補③', 'ステータス']);
      // ヘッダー行を太字に
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    }

    // データを追記
    sheet.appendRow([
      new Date().toLocaleString('ja-JP'),
      data.name,
      data.line,
      data.slot1,
      data.slot2,
      data.slot3,
      '未対応'
    ]);

    // メール通知（必要に応じてアドレスを変更）
    // MailApp.sendEmail({
    //   to: 'your-email@example.com',
    //   subject: '【個別戦略設計会】新規予約: ' + data.name,
    //   body: 'お名前: ' + data.name + '\n'
    //        + 'LINE名: ' + data.line + '\n'
    //        + '候補①: ' + data.slot1 + '\n'
    //        + '候補②: ' + data.slot2 + '\n'
    //        + '候補③: ' + data.slot3
    // });

    return ContentService
      .createTextOutput(JSON.stringify({status: 'ok'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// GETアクセス時（テスト用）
function doGet() {
  return ContentService
    .createTextOutput('予約システム稼働中')
    .setMimeType(ContentService.MimeType.TEXT);
}
