// ==========================================
// 【設定項目】
// ==========================================
const WEBHOOK_URL = '';
const SHEET_NAME  = '年間スケジュール_2026'; // 読み取りたいタブ名

// メンションしたい個人のユーザーID
const MENTION_ID = '<@UXXXXXXXXXX>';

// ==========================================

/**
 * 毎日深夜に実行してその日のトリガーを設定する
 */
function setDailyTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    console.error("シート「" + SHEET_NAME + "」が見つかりません。");
    return;
  }

  const data = sheet.getDataRange().getValues();
  
  // 今日の日付を「YYYY/MM/DD」の文字列で取得（タイムゾーンのズレを防止）
  const timeZone = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy/MM/dd');

  // 既存の巡回トリガーをリセット
  deleteTrigger('sendPatrolMessage');

  for (let i = 1; i < data.length; i++) {
    // 日付として無効な行はスキップ
    if (!data[i][0] || isNaN(new Date(data[i][0]).getTime())) continue;

    const rowDateStr = Utilities.formatDate(new Date(data[i][0]), timeZone, 'yyyy/MM/dd');

    // 今日が「開室日(C列が1)」かチェック
    if (rowDateStr === todayStr && data[i][2] == 1) {
      
      const dayOfWeek = data[i][1]; // B列：曜日
      let timeRange = String(data[i][4]); // E列：開室時間
      
      // 全角ハイフンや波線などを半角ハイフンに統一（表記揺れ対策）
      timeRange = timeRange.replace(/[ー〜−]/g, '-');
      
      if (!timeRange.includes('-')) {
        console.warn(`${todayStr} のE列にハイフンが含まれていません。`);
        continue;
      }
      
      const [startTime, endTime] = timeRange.split('-');
      const [endHour, endMin] = endTime.split(':').map(Number);

      // --- 巡回リマインドの時間を判定 ---
      const patrolTime = new Date();
      if (endHour < 18 || (endHour === 18 && endMin === 0)) {
        patrolTime.setHours(16, 0, 0, 0);
      } else {
        if (['月', '水', '金'].includes(dayOfWeek)) {
          patrolTime.setHours(17, 20, 0, 0);
        } else {
          patrolTime.setHours(17, 0, 0, 0);
        }
      }
      createTimeTrigger('sendPatrolMessage', patrolTime);
      
      // ログ出力
      console.log(`トリガー設定完了: 巡回 ${patrolTime.getHours()}:${patrolTime.getMinutes().toString().padStart(2, '0')}`);
      
      break;
    }
  }
}

/**
 * メッセージ：巡回
 */
function sendPatrolMessage() {
  const text = `${MENTION_ID} message`;
  postToSlack(text);
}

// --- 以下、補助関数（変更なし） ---

function postToSlack(text) {
  const payload = { "text": text };
  UrlFetchApp.fetch(WEBHOOK_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });
}

function createTimeTrigger(funcName, date) {
  if (date.getTime() < new Date().getTime()) return;
  ScriptApp.newTrigger(funcName).timeBased().at(date).create();
}

function deleteTrigger(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}