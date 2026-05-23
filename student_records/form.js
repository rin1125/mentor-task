/**
 * 設定：共有ドライブにあるスプレッドシートのID
 */
const SPREADSHEET_ID = '';

/**
 * SlackのWebhook URL設定
 * 提供いただいたURLを反映しています
 */
const SLACK_WEBHOOK_URLS = {
  CHANNEL_A: '', // 中1用
  CHANNEL_B: '', // 中2,3用
  CHANNEL_C: ''  // 高1,2,3用
};

/**
 * Webアプリの表示設定
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('生徒対応履歴フォーム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * スプレッドシートの「list」シートから学年別の生徒データを取得
 */
function getStudentData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('list');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2 || lastCol < 1) return {};

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const studentData = {};
  
  for (let colIndex = 0; colIndex < lastCol; colIndex++) {
    let gradeName = values[0][colIndex];
    if (!gradeName) continue;

    // 全角数字を半角に変換して正規化
    gradeName = gradeName.toString().replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
    
    const students = [];
    for (let rowIndex = 1; rowIndex < lastRow; rowIndex++) {
      const name = values[rowIndex][colIndex];
      if (name && name !== "") students.push(name);
    }
    studentData[gradeName] = students;
  }
  return studentData;
}

/**
 * フォーム送信時の処理（シートへの保存 ＆ Slack通知の実行）
 */
function processForm(formObject) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('interaction');
    
    const timestamp = new Date();
    const rowData = [
      timestamp,
      formObject.mentorName,
      formObject.studentGrade,
      formObject.studentName,
      formObject.subject,
      formObject.instruction,
      formObject.qualitativeInfo
    ];
    
    // 1. スプレッドシートの「interaction」シートへ追記
    sheet.appendRow(rowData);
    
    // 2. Slackへの通知実行
    sendToSlack(formObject);
    
    return "送信が完了しました！";
    
  } catch (e) {
    console.error("エラーが発生しました: " + e.toString());
    return "エラーが発生しました: " + e.message;
  }
}

/**
 * 学年に応じてSlackに通知を飛ばす内部関数（テキスト形式）
 */
function sendToSlack(formObject) {
  const grade = formObject.studentGrade;
  let webhookUrl = "";

  // 学年による振り分けロジック
  if (grade === "中1") {
    webhookUrl = SLACK_WEBHOOK_URLS.CHANNEL_A;
  } else if (grade === "中2" || grade === "中3") {
    webhookUrl = SLACK_WEBHOOK_URLS.CHANNEL_B;
  } else if (grade.indexOf("高") !== -1) { 
    webhookUrl = SLACK_WEBHOOK_URLS.CHANNEL_C;
  }

  if (!webhookUrl) {
    console.warn("該当する学年のSlack通知先が見つかりませんでした: " + grade);
    return;
  }

  const messageLines = [
    `*メンター名：* ${formObject.mentorName}`,
    `*学年：* ${formObject.studentGrade}`,
    `*生徒名：* ${formObject.studentName}`,
    `*科目：* ${formObject.subject}`,
    `*指導内容：*`,
    formObject.instruction,
    `*定性情報：*`,
    formObject.qualitativeInfo
  ];

  // 配列を改行コード（\n）で綺麗に結合する
  const messageText = messageLines.join('\n');

  const payload = {
    text: messageText
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  // Slackへ送信実行
  UrlFetchApp.fetch(webhookUrl, options);
}