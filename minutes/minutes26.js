/**
 * リーダー会議議事録を作成し、指定フォルダへ移動する
 */
function createLeaderMeetingMinutesC() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('UnitC'); 
  const data = sheet.getDataRange().getValues();
  data.shift(); // ヘッダーをスキップ
  
  const now = new Date();
  const twoWeeksAgo = new Date(now.getTime() - (14 * 24 * 60 * 60 * 1000));
  const dateStr = Utilities.formatDate(now, "JST", "yyyyMMdd");
  
  let reportContext = "";
  let interactionRawData = []; 
  
  let visitors = {
    g1: [], g2: [], g3: [], total: []
  };

  // 1. 直近2週間のデータを抽出・集計
  data.forEach((row) => {
    const rowDate = new Date(row[0]);
    if (rowDate >= twoWeeksAgo && rowDate <= now) {
      reportContext += `--- 勤務報告 ---\n`;
      reportContext += `生徒対応記録数: ${row[7]}, 様子: ${row[8]}\n`;
      reportContext += `工夫: ${row[9]}, 課題: ${row[10]}, 連絡事項: ${row[11]}\n`;
      reportContext += `次回目標: メンター1(${row[16]}), メンター2(${row[18]}), メンター3(${row[20]})\n\n`;

      if (row[7] !== "") {
        interactionRawData.push(row[7]);
      }

      let countG1 = parseInt(row[4], 10) || 0;
      let countG2 = parseInt(row[5], 10) || 0;
      let countG3 = parseInt(row[6], 10) || 0;
      let countTotal = countG1 + countG2 + countG3;

      visitors.g1.push(countG1);
      visitors.g2.push(countG2);
      visitors.g3.push(countG3);
      visitors.total.push(countTotal);
    }
  });

  const calcStats = (arr) => {
    if (arr.length === 0) return { max: 0, min: 0, avg: "0.0" };
    const max = Math.max(...arr);
    const min = Math.min(...arr);
    const avg = (arr.reduce((a, b) => a + b, 0) / arr.length).toFixed(1);
    return { max, min, avg };
  };

  const statG1 = calcStats(visitors.g1);
  const statG2 = calcStats(visitors.g2);
  const statG3 = calcStats(visitors.g3);
  const statTotal = calcStats(visitors.total);

  // 2. OpenAI APIによる分析
  const llmAnalysis = callOpenAiApi(reportContext, interactionRawData);

  // 3. Googleドキュメントの作成（この時点ではマイドライブ直下に作成される）
  const docTitle = `リーダーmtg_UnitC_${dateStr}`;
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();

  // --- 議事録フォーマットの構築 ---
  body.appendParagraph(docTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`開催日: ${Utilities.formatDate(now, "JST", "yyyy/MM/dd")}`);
  body.appendHorizontalRule();

  body.appendParagraph("1. 振り返り").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("■ 達成度評価").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  const evalTable = [
    ["項目", "最高値", "最低値", "平均", "備考"],
    ["来室人数(高1)", statG1.max.toString(), statG1.min.toString(), statG1.avg.toString(), ""],
    ["来室人数(高2)", statG2.max.toString(), statG2.min.toString(), statG2.avg.toString(), ""],
    ["来室人数(高3)", statG3.max.toString(), statG3.min.toString(), statG3.avg.toString(), ""],
    ["来室人数(合計)", statTotal.max.toString(), statTotal.min.toString(), statTotal.avg.toString(), ""],
    ["生徒対応数", llmAnalysis.student_interaction.max, llmAnalysis.student_interaction.min, llmAnalysis.student_interaction.average, ""]
  ];
  const table = body.appendTable(evalTable);
  table.getRow(0).setAttributes({[DocumentApp.Attribute.BOLD]: true});

  body.appendParagraph("■ Good").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph(llmAnalysis.good);

  body.appendParagraph("■ More").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph(llmAnalysis.more);

  body.appendParagraph("■ Next").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph("・\n・");

  body.appendParagraph("2. その他の議題").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem("");

  body.appendParagraph("3. 決定事項").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("いつまでに・誰が・何をやるのかを明確にする。");
  const decisionTable = [
    ["When (期限)", "Who (担当)", "What (タスク内容)"],
    ["", "", ""],
    ["", "", ""]
  ];
  const dTable = body.appendTable(decisionTable);
  dTable.getRow(0).setAttributes({[DocumentApp.Attribute.BOLD]: true});

  // --- 保存と移動の処理 ---

  // 1. ドキュメントを一旦保存して閉じる（競合防止）
  doc.saveAndClose();

  // 2. サーバー側の反映を待つための待機（Service Error対策）
  Utilities.sleep(1500);

  // 3. 指定のフォルダに移動
  const folderId = ""; // 例: "1a2b3c4d5e6f7g8h9i0j"
  try {
    const file = DriveApp.getFileById(doc.getId());
    const folder = DriveApp.getFolderById(folderId);
    file.moveTo(folder);
  } catch (e) {
    Logger.log("フォルダ移動中にエラーが発生しました: " + e.message);
    SpreadsheetApp.getUi().alert("ファイルは作成されましたが、フォルダ移動に失敗しました。IDを確認してください。");
    return;
  }

  const url = doc.getUrl();
  Logger.log(url);
}

/**
 * OpenAI API 呼び出し部分（変更なし）
 * ※ OPENAI_API_KEY は別途定義されている前提です
 */
function callOpenAiApi(reportText, interactionRawData) {
  const url = "https://api.openai.com/v1/chat/completions";
  const prompt = `以下の「メンター勤務ログ」を分析し、リーダー会議用に「良かった点(Good)」と「改善点(More)」を要点ごとにまとめ、3~4つ箇条書きで抽出してください。
さらに、「生徒対応記録数 生データ」には全角・半角や文字が混在しています。ここから正確に数値を抽出し、最高値、最低値、平均値（小数第1位まで）を計算してください。

必ず以下の構造のJSON形式で出力してください：
{
  "good": "抽出された良い点の文字列(\n・で区切る)",
  "more": "抽出された改善が必要な点の文字列(\n・で区切る)",
  "student_interaction": {
    "max": "最高値",
    "min": "最低値",
    "average": "平均値"
  }
}

【メンター勤務ログ】
${reportText}

【生徒対応記録数 生データ】
${interactionRawData.join("\n")}`;

  const payload = {
    "model": "gpt-4o-mini",
    "response_format": { "type": "json_object" },
    "messages": [
      { "role": "system", "content": "あなたは教育プログラムの分析リーダーです。" },
      { "role": "user", "content": prompt }
    ],
    "temperature": 0.7
  };

  const options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + "OPENAI_API_KEY",
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  const resContent = JSON.parse(response.getContentText());
  const analysis = JSON.parse(resContent.choices[0].message.content);

  return {
    good: analysis.good || "（データなし）",
    more: analysis.more || "（データなし）",
    student_interaction: analysis.student_interaction || { max: "-", min: "-", average: "-" }
  };
}

function createMtgTriggers() {
  // 会議の開催日リスト
  const mtgDates = [
    "2026/04/23", "2026/05/07", "2026/05/21", "2026/06/04",
    "2026/06/18", "2026/07/02", "2026/07/16"
  ];
  const startTime = "20:30"; // 会議開始時間
  const leadMinutes = 60;      // 何分前に実行するか（例：5分前）

  // 既存の同一関数のトリガーを一度削除（重複防止のため）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createLeaderMeetingMinutes') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  mtgDates.forEach(dateStr => {
    // 日付と時間を結合してDateオブジェクトを作成
    const triggerTime = new Date(`${dateStr} ${startTime}:00`);
    
    // 指定した分だけ時間を戻す（「直前」の設定）
    triggerTime.setMinutes(triggerTime.getMinutes() - leadMinutes);

    // すでに過ぎてしまった時間はスキップ
    if (triggerTime > new Date()) {
      ScriptApp.newTrigger('createLeaderMeetingMinutes') // 実行したい関数名
               .timeBased()
               .at(triggerTime)
               .create();
      
      console.log(`${dateStr} の ${leadMinutes}分前（${Utilities.formatDate(triggerTime, "JST", "HH:mm")}）にトリガーを設定しました。`);
    }
  });
}