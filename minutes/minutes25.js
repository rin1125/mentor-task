function generateUserReport() {
  const templateUrl = 'https://docs.google.com/document/d/1a2b3c4d5e6f7g8h9i0j/edit'; // テンプレートドキュメントのURL
  const templateId = templateUrl.match(/[-\w]{25,}/)[0];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("来室人数");
  const data = sheet.getDataRange().getValues();

  const copy = DriveApp.getFileById(templateId).makeCopy(`リーダーMTG_議事録_${new Date().toLocaleDateString()}`);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  const text = body.getText();

  // 日付取得
  const startMatch = text.match(/{{start_date}}:\s*(\d{4}\/\d{1,2}\/\d{1,2})/);
  const endMatch = text.match(/{{end_date}}:\s*(\d{4}\/\d{1,2}\/\d{1,2})/);
  if (!startMatch || !endMatch) {
    Logger.log("テンプレートに{{start_date}}または{{end_date}}が見つかりません");
    return;
  }
  const startDate = new Date(startMatch[1]);
  const endDate = new Date(endMatch[1]);

  // 利用人数データ取得（列指定）
  const userCols = [2, 3, 4, 5, 6, 7]; // 中1〜高3の列番号（0始まり）
  const headers = ["中1", "中2", "中3", "高1", "高2", "高3"];

  // 目標平均人数抽出
  const targetMatch = text.match(/【目標平均人数】([\s\S]+?)(?:\n\n|\Z)/);
  const targets = {};
  if (targetMatch) {
    const lines = targetMatch[1].trim().split('\n');
    lines.forEach(line => {
      const m = line.match(/(中\d|高\d):\s*(\d+)/);
      if (m) targets[m[1]] = parseInt(m[2]);
    });
  }

  // データ集計
  const dateList = [];
  const gradeData = headers.map(() => []);
  const gradeSums = Array(headers.length).fill(0);
  const gradeCounts = Array(headers.length).fill(0);
  const dailyTotals = [];
  let totalAll = 0;
  let countDays = 0;

  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]);
    if (rowDate >= startDate && rowDate <= endDate) {
      const rowVals = userCols.map(c => Number(data[i][c]));
      const validVals = rowVals.filter(v => !isNaN(v));
      const sum = validVals.reduce((a, b) => a + b, 0);
      if (sum === 0) continue;

      const dow = ["日", "月", "火", "水", "木", "金", "土"][rowDate.getDay()];
      const dateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "M/d") + `(${dow})`;
      dateList.push(dateStr);
      dailyTotals.push(sum);
      countDays++;

      for (let j = 0; j < headers.length; j++) {
        const val = Number(data[i][userCols[j]]);
        if (!isNaN(val)) {
          const rounded = Math.round(val);
          gradeData[j].push(rounded);
          gradeSums[j] += rounded;
          gradeCounts[j]++;
          totalAll += rounded;
        } else {
          gradeData[j].push("-");
        }
      }
    }
  }

  const avgTotal = countDays > 0 ? Math.round(totalAll / countDays) : 0;

  // プレースホルダ置換
  body.replaceText('{{start_date}}', startMatch[1]);
  body.replaceText('{{end_date}}', endMatch[1]);
  body.replaceText('{{average_users}}', String(avgTotal));

  // 表データ構築
  const tableData = [];
  const headerRow = ["学年"].concat(dateList).concat(["平均", "達成率"]);
  tableData.push(headerRow);

  for (let i = 0; i < headers.length; i++) {
    const row = [headers[i]];
    gradeData[i].forEach(val => row.push(String(val)));
    const avg = gradeCounts[i] > 0 ? Math.round(gradeSums[i] / gradeCounts[i]) : 0;
    const target = targets[headers[i]] || 1;
    const ratio = target > 0 ? Math.round((avg / target) * 100) : 0;
    row.push(String(avg));
    row.push(`${ratio}%`);
    tableData.push(row);
  }

  const totalRow = ["全体"];
  const overallTarget = headers.reduce((sum, h) => sum + (targets[h] || 0), 0);
  dailyTotals.forEach(val => totalRow.push(String(val)));
  totalRow.push(String(avgTotal));
  totalRow.push(overallTarget > 0 ? `${Math.round((avgTotal / overallTarget) * 100)}%` : "-");
  tableData.push(totalRow);

  // 表の挿入
  const placeholder = "{{USAGE_TABLE}}";
  const found = body.findText(placeholder);
  if (found) {
    const element = found.getElement();
    const paragraph = element.getParent();
    paragraph.clear();
    body.insertParagraph(body.getChildIndex(paragraph), "■学年別利用人数").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.insertTable(body.getChildIndex(paragraph) + 1, tableData);
  } else {
    body.appendParagraph("■学年別利用人数").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendTable(tableData);
  }
  const validGradesByUnit = {
  "UnitA生徒指導履歴": ["中学1年生"],
  "UnitB生徒指導履歴": ["中学２年生", "中学３年生"],
  "UnitC生徒指導履歴": ["高校１年生", "高校２年生", "高校３年生"]
};

const unitGradeCounts = [];

for (const sheetName in validGradesByUnit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) continue;

  const data = sheet.getDataRange().getValues();
  const targetGrades = validGradesByUnit[sheetName];

  for (const grade of targetGrades) {
    const count = data.filter(row => {
      const ts = row[0];
      const g = row[2];
      const date = new Date(ts);
      return g === grade && date >= startDate && date <= endDate;
    }).length;
    unitGradeCounts.push([sheetName.replace("生徒指導履歴", ""), grade, count]);
  }
}
const unitPlaceholder = '{{UNIT_COUNTS}}';
const unitFound = body.findText(unitPlaceholder);
const unitTableData = [['Unit', '学年', '指導件数']].concat(unitGradeCounts);

if (unitFound) {
  const unitElement = unitFound.getElement();
  const unitParagraph = unitElement.getParent();
  unitParagraph.clear();
  body.insertParagraph(body.getChildIndex(unitParagraph), '■学年別指導件数').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.insertTable(body.getChildIndex(unitParagraph) + 1, unitTableData);
} else {
  body.appendParagraph('■学年別指導件数').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendTable(unitTableData);
}


  doc.saveAndClose();
  Logger.log("✅ 完成ドキュメントURL: " + doc.getUrl());
}
