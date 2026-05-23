/**
 * Google Drive にフォルダ・スプレッドシートを自動生成するスクリプト
 *
 * 構成:
 *   マイドライブ（または指定フォルダ）/
 *     {学年}/                 ← 学年フォルダ（例: j1, j2）
 *       {学年}_{クラス}.xlsx  ← クラスごとのスプレッドシート（例: j1_A）
 *         タブ: "01 阿部 太陽", "02 安藤 紬", ...
 *
 * 使い方:
 *   1. master.csv（実体は .xlsx）を Google ドライブにアップロードし、
 *      「Googleスプレッドシートとして開く」で変換する
 *   2. そのスプレッドシートの URL から ID をコピーし、
 *      下の MASTER_SPREADSHEET_ID に貼り付ける
 *   3. 必要に応じて PARENT_FOLDER_ID・SHEET_NAME を変更する
 *   4. createFoldersAndSheets() を実行する
 */

// =========================================================
// ★ 設定項目 ★
// =========================================================

/** マスタースプレッドシートの ID（URL の /d/XXXX/edit の XXXX 部分） */
const MASTER_SPREADSHEET_ID = '1a2b3c4d5e6f7g8h9i0j'; // ← ここにマスタースプレッドシートのIDを貼り付けてください

/** ログ管理スプレッドシートの URL（IMPORTRANGE で使用） */
const MASTER_LOG_URL = 'https://docs.google.com/spreadsheets/d/1a2b3c4d5e6f7g8h9i0j/edit'; // ← ここにマスターログスプレッドシートのURLを貼り付けてください

/** 作成先の親フォルダ ID（空文字の場合はマイドライブ直下に作成） */
const PARENT_FOLDER_ID = '1a2b3c4d5e6f7g8h9i0j'; // ← ここに作成先の親フォルダIDを貼り付けてください（例: "1a2b3c4d5e6f7g8h9i0j"）

/** マスターシートのシート名 */
const MASTER_SHEET_NAME = 'master';

/**
 * 各クラスのスプレッドシート内タブの名前フォーマット
 * {num} → 出席番号（2桁ゼロ埋め）, {name} → 氏名
 */
const TAB_NAME_FORMAT = '{num}_{name}';

/** 名簿シートのタブ名 */
const SUMMARY_SHEET_NAME = '生徒様一覧';

// =========================================================
// メイン処理
// =========================================================

function createFoldersAndSheets() {
  // ----- マスターデータ読み込み -----
  const masterSS = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const masterSheet = masterSS.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    throw new Error(`シート "${MASTER_SHEET_NAME}" が見つかりません`);
  }

  const data = masterSheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('マスターデータが空です');
  }

  // ヘッダー: 学年, クラス, 出席番号, 生徒番号, 氏名
  const rows = data.slice(1);

  // ----- 学年×クラス → 生徒リスト のマップ作成 -----
  /** @type {Object.<string, Object.<string, Array.<{num:number, name:string}>>>} */
  const classMap = {};

  for (const row of rows) {
    const grade     = String(row[0]).trim();
    const cls       = String(row[1]).trim();
    const num       = Number(row[2]);
    const studentId = String(row[3]).trim();
    const name      = String(row[4]).trim();

    if (!grade || !cls || isNaN(num) || !name) continue;

    classMap[grade]       = classMap[grade]       || {};
    classMap[grade][cls]  = classMap[grade][cls]  || [];
    classMap[grade][cls].push({ grade, cls, num, studentId, name });
  }

  // ----- 親フォルダ取得 -----
  // URL が渡された場合はフォルダ ID を抽出する
  // 例: https://drive.google.com/drive/folders/XXXX → XXXX
  const parentFolderId = PARENT_FOLDER_ID
    ? PARENT_FOLDER_ID.replace(/^.*\/folders\/([^/?]+).*$/, '$1')
    : '';
  const parentFolder = parentFolderId
    ? DriveApp.getFolderById(parentFolderId)
    : DriveApp.getRootFolder();

  // ----- 学年フォルダ → クラス別スプレッドシート → タブ作成 -----
  const grades = Object.keys(classMap).sort();

  for (const grade of grades) {
    // 学年フォルダ（既存なら再利用）
    const gradeFolder = getOrCreateFolder(parentFolder, grade);
    Logger.log(`📂 フォルダ: ${grade}`);

    const classes = Object.keys(classMap[grade]).sort();

    for (const cls of classes) {
      // 出席番号順にソート
      const students = classMap[grade][cls].sort((a, b) => a.num - b.num);
      const ssName   = `${grade}_${cls}`;

      // スプレッドシート（既存なら再利用）
      const ss = getOrCreateSpreadsheet(gradeFolder, ssName);
      const existingSheetNames = new Set(ss.getSheets().map(s => s.getName()));

      if (existingSheetNames.has(SUMMARY_SHEET_NAME)) {
        Logger.log(`  ✅ スキップ（完了済み）: ${ssName}`);
        continue;
      }

      Logger.log(`  📊 処理中: ${ssName}`);

      // IMPORTRANGE の事前認証（ブラウザでの手動クリックを不要にする）
      authorizeImportRange(ss.getId(), MASTER_SPREADSHEET_ID);

      setupSheetTabs(ss, students);
      SpreadsheetApp.flush();          // タブをサーバーへコミット（参照先を確定）
      setupSummarySheet(ss, students); // タブが確定した後に数式を書き込む

      // 全操作を確定してから次のスプレッドシートへ
      SpreadsheetApp.flush();
    }
  }

  Logger.log('✅ 全処理完了');
}

// =========================================================
// 名簿シートの作成・更新
// =========================================================

/**
 * 各クラスの生徒情報一覧を「名簿」シートに書き出す
 * 既存の場合は内容を上書きする
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Array.<{grade:string, cls:string, num:number, studentId:string, name:string}>} students
 */
function setupSummarySheet(ss, students) {
  // 既存なら再実行不要（タイムアウト後の再開時もスキップ）
  if (ss.getSheetByName(SUMMARY_SHEET_NAME)) {
    Logger.log(`    → 名簿シートは既存のためスキップ`);
    return;
  }

  // 新規スプレッドシートのデフォルトシートをリネームして再利用
  // ※ SpreadsheetApp.create() 直後に insertSheet() を呼ぶと
  //   "Sheet 0 not found" が発生するため、リネームで回避する
  // ※ setupSheetTabs が先に実行済みのため、getSheets()[0] は
  //   insertSheet で末尾追加されている学生タブではなく、
  //   元のデフォルトシート（シート1）になる
  const sheet = ss.getSheets()[0];
  sheet.setName(SUMMARY_SHEET_NAME);
  SpreadsheetApp.flush();
  Logger.log(`    ＋ 名簿シート作成: ${SUMMARY_SHEET_NAME}`);

  // ヘッダー行（基本5列＋集計4列）
  const headers = ['学年', 'クラス', '出席番号', '生徒番号', '氏名', '来室回数', '来室時間', '生徒対応履歴数', 'イベント参加回数'];
  const totalCols = headers.length;
  sheet.getRange(1, 1, 1, totalCols).setValues([headers]);

  // ヘッダー書式（太字・背景色）
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // 生徒データ行（基本5列）
  const dataRows = students.map(({ grade, cls, num, studentId, name }) =>
    [grade, cls, num, studentId, name]
  );
  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, 5).setValues(dataRows);

    // 集計数式を各行に設定
    // 同じスプレッドシート内の各生徒タブから直接読み取る
    for (let i = 0; i < students.length; i++) {
      const row = i + 2;
      const student = students[i];
      const tabName = TAB_NAME_FORMAT
        .replace('{num}',  String(student.num).padStart(2, '0'))
        .replace('{name}', student.name);

      // F列: 来室回数 → 各生徒タブのB16（来室回数合計）
      const visitCountFormula = `='${tabName}'!B16`;

      // G列: 来室時間 → 各生徒タブのC16（合計時間）
      const visitTimeFormula = `='${tabName}'!C16`;

      // H列: 生徒対応履歴数 → 各生徒タブのA20以降のデータ行数（ヘッダー除く）
      const counselingCountFormula = `=IFERROR(COUNTA('${tabName}'!A20:A200)-COUNTIF('${tabName}'!A20:A200,"（履歴なし）"),0)`;

      // I列: イベント参加回数 → 各生徒タブのE4以降のデータ行数（ヘッダー除く）
      const eventCountFormula = `=IFERROR(COUNTA('${tabName}'!E4:E200)-COUNTIF('${tabName}'!E4:E200,"（参加履歴なし）"),0)`;

      sheet.getRange(row, 6).setFormula(visitCountFormula);
      sheet.getRange(row, 7).setFormula(visitTimeFormula);
      sheet.getRange(row, 7).setNumberFormat('[h]:mm:ss');
      sheet.getRange(row, 8).setFormula(counselingCountFormula);
      sheet.getRange(row, 9).setFormula(eventCountFormula);
    }

    // データ行の書式（1行おきに色分け）
    for (let i = 0; i < students.length; i++) {
      const row = i + 2;
      const rowRange = sheet.getRange(row, 1, 1, totalCols);
      rowRange.setBackground(row % 2 === 0 ? '#ffffff' : '#f0f4ff');
      rowRange.setVerticalAlignment('middle');
      sheet.setRowHeight(row, 28);
    }

    // 罫線
    sheet.getRange(1, 1, students.length + 1, totalCols).setBorder(
      true, true, true, true, true, true,
      '#a0b4e8', SpreadsheetApp.BorderStyle.SOLID
    );
    sheet.getRange(1, 1, 1, totalCols).setBorder(
      true, true, true, true, null, null,
      '#1a73e8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  }

  // 列幅設定
  sheet.setColumnWidth(1, 55);   // 学年
  sheet.setColumnWidth(2, 60);   // クラス
  sheet.setColumnWidth(3, 75);   // 出席番号
  sheet.setColumnWidth(4, 100);  // 生徒番号
  sheet.setColumnWidth(5, 130);  // 氏名
  sheet.setColumnWidth(6, 80);   // 来室回数
  sheet.setColumnWidth(7, 100);  // 来室時間
  sheet.setColumnWidth(8, 130);  // 生徒対応履歴数
  sheet.setColumnWidth(9, 130);  // イベント参加回数

}

// =========================================================
// タブ（シート）の整備
// =========================================================

/**
 * スプレッドシート内のタブを生徒リストに合わせて作成・リネーム・削除する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Array.<{num:number, name:string}>} students
 */
function setupSheetTabs(ss, students) {
  // A3: 月ラベル（4月〜翌3月の12行、ヘッダー含む）
  const monthLabelFormula =
    `=VSTACK("月",MAP(SEQUENCE(12,1,4),LAMBDA(col,IF(col<=12,col,col-12))))`;

  // B3: 月別入退室回数（SUMPRODUCT でカウント）
  const monthlyVisitCountFormula =
    `=LET(src,IMPORTRANGE("${MASTER_LOG_URL}","attendance!C:G"),` +
    `ids,CHOOSECOLS(src,1),` +
    `dates,CHOOSECOLS(src,3),` +
    `VSTACK("回数",MAP(SEQUENCE(12,1,4),LAMBDA(col,` +
    `LET(m,IF(col<=12,col,col-12),y,IF(m>=4,2026,2027),` +
    `SUMPRODUCT((ids=$D$1)*(dates>=DATE(y,m,1))*(dates<=EOMONTH(DATE(y,m,1),0))))))))`;

  // C3: 月別合計利用時間（SUMPRODUCT で合計）
  const monthlyUsageTimeFormula =
    `=LET(src,IMPORTRANGE("${MASTER_LOG_URL}","attendance!C:G"),` +
    `ids,CHOOSECOLS(src,1),` +
    `dates,CHOOSECOLS(src,3),` +
    `in_times,CHOOSECOLS(src,4),` +
    `out_times,CHOOSECOLS(src,5),` +
    `durations,MAP(in_times,out_times,LAMBDA(i,o,IF(OR(i="",o=""),0,IFERROR(MOD(VALUE(o)-VALUE(i),1),0)))),` +
    `VSTACK("合計時間",MAP(SEQUENCE(12,1,4),LAMBDA(col,` +
    `LET(m,IF(col<=12,col,col-12),y,IF(m>=4,2026,2027),` +
    `SUMPRODUCT((ids=$D$1)*(dates>=DATE(y,m,1))*(dates<=EOMONTH(DATE(y,m,1),0))*durations))))))`;

  // A18: 当該生徒の対応履歴（headers=1 で数式自体がヘッダー行を含む）
  const counselingHistoryFormula =
    `=IFERROR(QUERY(IMPORTRANGE("${MASTER_LOG_URL}","interaction!A:G"),` +
    `"SELECT Col1, Col5, Col6, Col7 WHERE Col4 = '"&D1&"'",1),` +
    `{"日時","教科","指導内容","定性情報";"（履歴なし）","","",""})`; 

  // E3: イベント履歴（event!シートのG列がD1と一致する行のB,C,H:M列を見出し付きで展開）
  const eventHistoryFormula =
    `=LET(` +
      `src,IMPORTRANGE("${MASTER_LOG_URL}","event!A:M"),` +
      `cols,CHOOSECOLS(src,2,3,8,9,10,11,12,13),` +
      `ids,CHOOSECOLS(src,7),` +
      `hdr,{"実施日","イベント名","満足度","満足度の理由","疑問解消","今後進めたい","感想","要望"},` +
      `body,IFERROR(FILTER(cols,ids=D1),{"（参加履歴なし）","","","","","","",""}),` +
      `VSTACK(hdr,body)` +
    `)`;

  for (const student of students) {
    const tabName = TAB_NAME_FORMAT
      .replace('{num}',  String(student.num).padStart(2, '0'))
      .replace('{name}', student.name);

    // タブが未作成なら末尾に追加
    // ※ insertSheet(name) はアクティブシートの後に挿入されるため、
    //   途中再開時に順番が崩れる。明示的に末尾インデックスを指定して回避する。
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName, ss.getSheets().length);
      sheet.getRange(1, 1, 1, 4).setValues([[student.grade, student.cls, student.num, student.name]]);
      Logger.log(`    ＋ タブ追加: ${tabName}`);
    }

    // A3 に数式が入っていれば書き込み完了済み → スキップ
    // （タイムアウトで途中停止した場合でも、未完了タブのみ再処理できる）
    if (sheet.getRange('A3').getFormula()) continue;
    Logger.log(`    ✏️ 数式・書式設定: ${tabName}`);

    sheet.getRange('A3').setFormula(monthLabelFormula);
    sheet.getRange('B3').setFormula(monthlyVisitCountFormula);
    sheet.getRange('C3').setFormula(monthlyUsageTimeFormula);
    sheet.getRange('C4:C15').setNumberFormat('[h]:mm:ss');
    sheet.getRange('A19').setFormula(counselingHistoryFormula);

    // ── 書式設定 ──────────────────────────────────────────────────

    // 1. 1行目（生徒情報）: 大きい文字・目立つ書式
    const row1 = sheet.getRange('A1:D1');
    row1.setFontSize(18);
    row1.setFontWeight('bold');
    row1.setBackground('#1a73e8');
    row1.setFontColor('#ffffff');
    row1.setHorizontalAlignment('center');
    row1.setVerticalAlignment('middle');
    sheet.setRowHeight(1, 48);
    sheet.setFrozenRows(1);

    // 2. 月別利用状況 ─ セクションタイトル（A2:C2）
    sheet.getRange('A2').setValue('月別利用状況');
    const sec1 = sheet.getRange('A2:C2');
    sec1.setFontSize(11);
    sec1.setFontWeight('bold');
    sec1.setFontColor('#1a73e8');
    sec1.setBackground('#e8f0fe');
    sec1.setVerticalAlignment('middle');
    sheet.setRowHeight(2, 28);

    // 3. 月別テーブル ─ ヘッダー行（A3:C3）とイベント数式（E3）
    const tblHead = sheet.getRange('A3:C3');
    tblHead.setFontWeight('bold');
    tblHead.setBackground('#4a86e8');
    tblHead.setFontColor('#ffffff');
    tblHead.setHorizontalAlignment('center');
    tblHead.setVerticalAlignment('middle');
    // E3：イベント履歴（ヘッダー行含む数式）
    sheet.getRange('E3').setFormula(eventHistoryFormula);
    const eventHead = sheet.getRange('E3:L3');
    eventHead.setFontWeight('bold');
    eventHead.setBackground('#4a86e8');
    eventHead.setFontColor('#ffffff');
    eventHead.setHorizontalAlignment('center');
    eventHead.setVerticalAlignment('middle');
    // E列（実施日）: 日付表示形式
    sheet.getRange('E4:E200').setNumberFormat('M/d');
    sheet.setRowHeight(3, 28);

    // 4. 月別テーブル ─ データ行（A4:C15）: 1行おきに色分け
    for (let r = 4; r <= 15; r++) {
      const rowRange = sheet.getRange(r, 1, 1, 3);
      rowRange.setBackground(r % 2 === 0 ? '#ffffff' : '#f0f4ff');
      rowRange.setHorizontalAlignment('center');
      rowRange.setVerticalAlignment('middle');
    }

    // 5. 月別テーブル ─ 罫線（外枠: 中実線、内側: 細実線）
    sheet.getRange('A3:C15').setBorder(
      true, true, true, true, null, null,
      '#1a73e8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
    sheet.getRange('A3:C15').setBorder(
      null, null, null, null, true, true,
      '#a0b4e8', SpreadsheetApp.BorderStyle.SOLID
    );

    // 6. 対応履歴セクション（A18：タイトル、A19：QUERYヘッダー行）
    sheet.getRange('A18').setValue('生徒対応履歴');
    const sec2 = sheet.getRange('A18:D18');
    sec2.setFontSize(11);
    sec2.setFontWeight('bold');
    sec2.setFontColor('#1a73e8');
    sec2.setBackground('#e8f0fe');
    sec2.setVerticalAlignment('middle');

    // A19: QUERYが出力するヘッダー行を青色に
    const counselHead = sheet.getRange('A19:D19');
    counselHead.setFontWeight('bold');
    counselHead.setBackground('#4a86e8');
    counselHead.setFontColor('#ffffff');
    counselHead.setHorizontalAlignment('center');
    counselHead.setVerticalAlignment('middle');
    // A20以降の日付列: 年度表示なし（M/d 形式）
    sheet.getRange('A20:A200').setNumberFormat('M/d');

    // 8. イベント履歴（E3から数式開始、ヘッダー行含む）
    sheet.getRange('E2').setValue('イベント参加履歴');
    const sec3 = sheet.getRange('E2:L2');
    sec3.setFontSize(11);
    sec3.setFontWeight('bold');
    sec3.setFontColor('#1a73e8');
    sec3.setBackground('#e8f0fe');
    sec3.setVerticalAlignment('middle');

    // 月別テーブル合計行（A16:C16）
    sheet.getRange('A16').setValue('合計');
    sheet.getRange('B16').setFormula('=SUM(B4:B15)');
    sheet.getRange('C16').setFormula('=SUM(C4:C15)');
    sheet.getRange('C16').setNumberFormat('[h]:mm:ss');
    const totalRow = sheet.getRange('A16:C16');
    totalRow.setFontWeight('bold');
    totalRow.setBackground('#c9daf8');
    totalRow.setHorizontalAlignment('center');
    totalRow.setVerticalAlignment('middle');
    sheet.setRowHeight(16, 28);

    // 10. 列幅設定
    sheet.setColumnWidth(1, 70);   // A: 学年 / 月
    sheet.setColumnWidth(2, 100);  // B: クラス / 回数
    sheet.setColumnWidth(3, 250);  // C: 出席番号 / 合計時間
    sheet.setColumnWidth(4, 250);  // D: 氏名 / 対応内容
    sheet.setColumnWidth(5, 100);  // E: イベント実施日
    sheet.setColumnWidth(6, 160);  // F: イベント名
    sheet.setColumnWidth(7, 80);   // G: 満足度
    sheet.setColumnWidth(8, 250);  // H: 満足度の理由
    sheet.setColumnWidth(9, 80);   // I: 疑問解消
    sheet.setColumnWidth(10, 250); // J: 今後進めたい
    sheet.setColumnWidth(11, 250); // K: 感想
    sheet.setColumnWidth(12, 250); // L: 要望

    // C・D列のデータ行: 折り返し表示
    sheet.getRange('C1:D200').setWrap(true);
    sheet.getRange('H1:H200').setWrap(true);
    sheet.getRange('J1:L200').setWrap(true);
    // 対応履歴データ行（A20以降）: 行高さを広めに
    for (let r = 20; r <= 100; r++) {
      sheet.setRowHeight(r, 60);
    }
  }
}

// =========================================================
// ユーティリティ
// =========================================================

/**
 * フォルダを取得または作成する
 * @param {GoogleAppsScript.Drive.Folder} parent
 * @param {string} name
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

/**
 * スプレッドシートを取得または作成し、指定フォルダに配置する
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getOrCreateSpreadsheet(folder, name) {
  // フォルダ内の同名ファイルを検索
  const it = folder.getFilesByName(name);
  if (it.hasNext()) {
    return SpreadsheetApp.open(it.next());
  }

  // 新規作成（ルートまたは任意の場所に生成される）
  const ss   = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(ss.getId());

  // 指定フォルダへ追加
  folder.addFile(file);

  // 他のすべての親フォルダから除外（ルート固定でなく汎用的に対応）
  const parents = file.getParents();
  while (parents.hasNext()) {
    const parent = parents.next();
    if (parent.getId() !== folder.getId()) {
      parent.removeFile(file);
    }
  }

  return ss;
}

/**
 * 宛先スプレッドシートから参照元スプレッドシートへの
 * IMPORTRANGE を事前認証する（ブラウザの手動クリック不要）
 *
 * Google Sheets の内部エンドポイントを呼び出し、
 * (宛先 SS, 参照元 SS) ペアの権限を付与する。
 * 既に認証済みの場合は何もしない（冪等）。
 *
 * @param {string} destSpreadsheetId  - 認証を付与したい宛先スプレッドシートの ID
 * @param {string} sourceSpreadsheetId - 参照元スプレッドシートの ID
 */
function authorizeImportRange(destSpreadsheetId, sourceSpreadsheetId) {
  const url =
    `https://docs.google.com/spreadsheets/d/${destSpreadsheetId}` +
    `/externaldata/addimportrangepermissions?donorDocId=${sourceSpreadsheetId}`;
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() !== 200) {
    Logger.log(`  ⚠️ IMPORTRANGE 認証 失敗 (${response.getResponseCode()}): ${destSpreadsheetId}`);
  } else {
    Logger.log(`  🔓 IMPORTRANGE 認証 完了: ${destSpreadsheetId}`);
  }
}
