/**
 * KNAemailsender - Kumon Email Sender
 *
 * Auto-synced from GitHub. Import folder: CSVs dropped into a Drive folder
 * are imported to "Math Data" and "Reading Data" tabs, then archived.
 */

// --- Version (bump when you deploy changes) ---
const VERSION = '1.0.25';

// --- Import folder config ---
const IMPORT_FOLDER_NAME = 'KNA Email Sender Import';
const ARCHIVE_FOLDER_NAME = 'Archive';

/**
 * Runs when the spreadsheet is opened. Adds a custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KNA Email Sender (v' + VERSION + ')')
    .addItem('Import from Drive', 'importFromDrive')
    .addItem('Create / Open Import Folder', 'createOrOpenImportFolder')
    .addSeparator()
    .addItem('Load', 'loadToWorkArea')
    .addItem('Move', 'syncDashboardToLog')
    .addToUi();
}

/**
 * Creates the import folder and Archive subfolder in Drive (if missing),
 * then opens the import folder in a new tab.
 */
function createOrOpenImportFolder() {
  const folder = getOrCreateImportFolder();
  if (folder) {
    SpreadsheetApp.getUi().alert(
      'Import folder ready',
      'Open this folder and drop your Math and Reading CSV files here:\n\n' + folder.getUrl(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Gets the import folder, creating it and the Archive subfolder if they don't exist.
 * @returns {GoogleAppsScript.Drive.Folder} The import folder, or null on error
 */
function getOrCreateImportFolder() {
  const root = DriveApp.getRootFolder();
  let importFolder = getChildFolderByName(root, IMPORT_FOLDER_NAME);
  if (!importFolder) {
    importFolder = root.createFolder(IMPORT_FOLDER_NAME);
  }
  let archiveFolder = getChildFolderByName(importFolder, ARCHIVE_FOLDER_NAME);
  if (!archiveFolder) {
    archiveFolder = importFolder.createFolder(ARCHIVE_FOLDER_NAME);
  }
  return importFolder;
}

/**
 * Finds a direct child folder by name.
 * @param {GoogleAppsScript.Drive.Folder} parent
 * @param {string} name
 * @returns {GoogleAppsScript.Drive.Folder|null}
 */
function getChildFolderByName(parent, name) {
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : null;
}

/**
 * Finds the most recently modified CSV in the folder whose name contains the subject (e.g. "Math" or "Reading").
 * Expects filenames like: Study Analysis_US WEST VIRTUAL_Math_Feb 2026_Feb 2026_02042026.csv
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @param {string} subject - "Math" or "Reading"
 * @returns {GoogleAppsScript.Drive.File|null}
 */
function findLatestCsvBySubject(folder, subject) {
  const files = folder.getFiles();
  let latest = null;
  let latestDate = 0;
  const subjectLower = subject.toLowerCase();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === 'text/csv' && file.getName().toLowerCase().indexOf(subjectLower) !== -1) {
      const date = file.getLastUpdated().getTime();
      if (date > latestDate) {
        latestDate = date;
        latest = file;
      }
    }
  }
  return latest;
}

/**
 * Imports a Drive CSV file into a sheet. Wipes any existing data and replaces with the CSV.
 * @param {GoogleAppsScript.Drive.File} csvFile
 * @param {string} sheetName - "Math Data" or "Reading Data"
 */
function importCsvToSheet(csvFile, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  const csvBlob = csvFile.getBlob();
  const csvText = csvBlob.getDataAsString();
  const rows = Utilities.parseCsv(csvText);
  if (rows.length === 0) return;
  // Wipe entire sheet (content + formatting) and replace with imported data
  sheet.clear();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.autoResizeColumns(1, rows[0].length);
}

/**
 * Moves a Drive file into the archive folder (add to archive, remove from current parent).
 * @param {GoogleAppsScript.Drive.File} file
 * @param {GoogleAppsScript.Drive.Folder} archiveFolder
 */
function moveToArchive(file, archiveFolder) {
  const parents = file.getParents();
  if (parents.hasNext()) {
    parents.next().removeFile(file);
  }
  archiveFolder.addFile(file);
}

/**
 * Main import: reads Math and Reading CSVs from the import folder,
 * writes them to "Math Data" and "Reading Data", then moves the files to Archive.
 */
function importFromDrive() {
  const ui = SpreadsheetApp.getUi();
  const importFolder = getOrCreateImportFolder();
  const archiveFolder = getChildFolderByName(importFolder, ARCHIVE_FOLDER_NAME);

  const mathFile = findLatestCsvBySubject(importFolder, 'Math');
  const readingFile = findLatestCsvBySubject(importFolder, 'Reading');

  const messages = [];

  if (mathFile) {
    try {
      importCsvToSheet(mathFile, 'Math Data');
      moveToArchive(mathFile, archiveFolder);
      messages.push('Math: imported and archived "' + mathFile.getName() + '"');
    } catch (e) {
      messages.push('Math: error - ' + e.message);
    }
  } else {
    messages.push('Math: no CSV file found in import folder');
  }

  if (readingFile) {
    try {
      importCsvToSheet(readingFile, 'Reading Data');
      moveToArchive(readingFile, archiveFolder);
      messages.push('Reading: imported and archived "' + readingFile.getName() + '"');
    } catch (e) {
      messages.push('Reading: error - ' + e.message);
    }
  } else {
    messages.push('Reading: no CSV file found in import folder');
  }

  ui.alert('Import from Drive', messages.join('\n'), ui.ButtonSet.OK);
}

// --- Dashboard → Log sync (button entry point) ---
// Log: A–C = Math Sent, E–G = Reading Sent, I–N = Issue (Subject,LoginID,Name,Trigger #,Note,Date)
//
// FILTER: To have Status and Notes shift up with the filtered rows, include M and N in the FILTER array.
// Then the spilled result has Status and Notes as the last columns (correct for each row). Those cells are
// read-only (formula output); to edit Status/Notes, edit the source row on the dashboard.
//
// Math (A,B,F,G + M,N so Status/Notes align):
//   =FILTER({'Math Dashboard'!A:A,'Math Dashboard'!B:B,'Math Dashboard'!F:F,'Math Dashboard'!G:G,'Math Dashboard'!M:M,'Math Dashboard'!N:N}, 'Math Dashboard'!E:E="SEND EMAIL", ISNA(MATCH('Math Dashboard'!A:A, FILTER(Log!J:J, Log!I:I="Math", Log!N:N=TODAY()), 0)))
//
// Reading: same pattern with 'Reading Dashboard'! and Log!I:I="Reading".

/** Expected dashboard headers (row 1). Status values: Not Sent, Issue, Sent. */
const DASHBOARD_HEADERS = ['LoginID', 'Name', 'Trigger #', 'Email', 'Status', 'Notes'];

/**
 * Syncs the current dashboard to the Log sheet. Tie this function to your button.
 * Runs on whichever sheet is active (must be "Math Dashboard" or "Reading Dashboard").
 * - Ignores "Not Sent".
 * - "Issue": appends Subject, LoginID, Name, Trigger #, Note to Log columns I–M (Subject in I).
 * - "Sent": appends LoginID, Name, Trigger # to Log columns A–C (Math) or E–G (Reading).
 */
function syncDashboardToLog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const name = sheet.getName();
    debugLog('Move', 'syncDashboardToLog start', { sheet: name });

    const isMath = name.toLowerCase().indexOf('math') !== -1;
    const isReading = name.toLowerCase().indexOf('reading') !== -1;
    const isDashboard = name.toLowerCase().indexOf('dashboard') !== -1;

    if (!isDashboard || (!isMath && !isReading)) {
      SpreadsheetApp.getUi().alert(
        'Wrong sheet',
        'Please run this from "Math Dashboard" or "Reading Dashboard".',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    const subject = isMath ? 'Math' : 'Reading';
    let logSheet = ss.getSheetByName('Log');
  if (!logSheet) {
    logSheet = ss.insertSheet('Log');
    // Set headers for Log (A–C Math Sent, E–G Reading Sent, I–N Issue). getRange(row, col, numRows, numCols)
    const h1 = [['Math LoginID', 'Math Name', 'Math Trigger #']];
    const h2 = [['Reading LoginID', 'Reading Name', 'Reading Trigger #']];
    const h3 = [['Subject', 'LoginID', 'Name', 'Trigger #', 'Note', 'Date']];
    logSheet.getRange(1, 1, 1, 3).setValues(h1);
    logSheet.getRange(1, 5, 1, 3).setValues(h2);
    logSheet.getRange(1, 9, 1, 6).setValues(h3);
  }

  // I–N headers are in row 2; data starts row 3
  const headerRow = sheet.getRange(2, 1, 2, sheet.getLastColumn()).getValues()[0];
  const col = getColumnIndices(headerRow);
  // Use the LoginID/Name/Trigger # block that contains Status (second block: I,J,K,L,M,N = LoginID,Name,Email,Trigger #,Status,Notes)
  if (col.status >= 4) {
    col.loginId = col.status - 4;
    col.name = col.status - 3;
    col.triggerNum = col.status - 1; // Trigger # is in L (Email is K = status - 2)
  }

  if (!col.loginId || !col.status) {
    SpreadsheetApp.getUi().alert(
      'Missing headers',
      'Dashboard must have headers for I–N (LoginID, Name, Trigger #, Email, Status, Notes) in row 2.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const issueRows = [];
  const issueSheetRows = [];
  const sentMathRows = [];
  const sentMathSheetRows = [];
  const sentReadingRows = [];
  const sentReadingSheetRows = [];

  const dataRange = sheet.getRange(3, 1, lastRow, sheet.getLastColumn());
  const rows = dataRange.getValues();

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const loginId = row[col.loginId - 1];
    // Only process rows that have work-area data (LoginID in I); skip blank rows
    if (loginId === null || loginId === undefined || String(loginId).trim() === '') continue;

    const status = String((row[col.status - 1] || '')).trim();
    if (status === '' || status.toLowerCase() === 'not sent') continue;

    const studentName = row[col.name - 1];
    const triggerNum = row[col.triggerNum - 1];
    const note = (col.notes && row[col.notes - 1] != null) ? row[col.notes - 1] : '';

    if (status.toLowerCase() === 'issue') {
      issueRows.push([subject, loginId, studentName, triggerNum, note]);
      issueSheetRows.push(3 + i);
    } else if (status.toLowerCase() === 'sent') {
      if (isMath) {
        sentMathRows.push([loginId, studentName, triggerNum]);
        sentMathSheetRows.push(3 + i);
      } else {
        sentReadingRows.push([loginId, studentName, triggerNum]);
        sentReadingSheetRows.push(3 + i);
      }
    }
  }

  // Append to Log. getRange(row, col, numRows, numCols) — 3rd/4th args are numRows/numCols, not endRow/endCol
  if (sentMathRows.length > 0) {
    const nextRow = getNextLogRow(logSheet, 1) + 1;
    const numRows = sentMathRows.length;
    logSheet.getRange(nextRow, 1, numRows, 3).setValues(sentMathRows);
  }
  if (sentReadingRows.length > 0) {
    const nextRow = getNextLogRow(logSheet, 5) + 1;
    const numRows = sentReadingRows.length;
    logSheet.getRange(nextRow, 5, numRows, 3).setValues(sentReadingRows);
  }
  if (issueRows.length > 0) {
    const nextRow = getNextLogRow(logSheet, 9) + 1;
    const numRows = issueRows.length;
    // Issue: I–N = Subject,LoginID,Name,Trigger #,Note,Date (do not write to G; G is for Reading Sent Trigger # only)
    logSheet.getRange(nextRow, 9, numRows, 5).setValues(issueRows);
    // Use spreadsheet's TODAY() so the date matches the sheet timezone (not script/server timezone)
    const todayCell = logSheet.getRange(1, 20);
    todayCell.setFormula('=TODAY()');
    SpreadsheetApp.flush();
    const today = todayCell.getValue();
    todayCell.clearContent();
    const dateCol = [];
    for (let i = 0; i < numRows; i++) {
      dateCol.push([today]);
    }
    logSheet.getRange(nextRow, 14, dateCol.length, 1).setValues(dateCol);
    // Clear Status and Notes on the dashboard for the rows we just logged
    for (let r = 0; r < issueSheetRows.length; r++) {
      const sheetRow = issueSheetRows[r];
      sheet.getRange(sheetRow, col.status).clearContent();
      if (col.notes) sheet.getRange(sheetRow, col.notes).clearContent();
    }
  }

  // Delete full rows for logged rows (from bottom to top)
  const rowsToDelete = issueSheetRows.concat(sentMathSheetRows || [], sentReadingSheetRows || []);
  if (rowsToDelete.length > 0) {
    const sorted = rowsToDelete.slice().sort(function (a, b) { return b - a; });
    for (let d = 0; d < sorted.length; d++) {
      sheet.deleteRow(sorted[d]);
    }
  }

    const msg = [
      subject + ' Dashboard → Log',
      'Sent: ' + (isMath ? sentMathRows.length : sentReadingRows.length),
      'Issues: ' + issueRows.length,
      rowsToDelete.length > 0 ? 'Rows removed from sheet.' : ''
    ].join('\n');
    debugLog('Move', 'complete', { sheet: name, sent: isMath ? sentMathRows.length : sentReadingRows.length, issues: issueRows.length });
    SpreadsheetApp.getUi().alert('Move complete', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    debugLog('Move', 'ERROR', { message: e.message, stack: e.stack });
    SpreadsheetApp.getUi().alert('Move error', e.message + '\n\nCheck View > Logs (Execution log) for details.', SpreadsheetApp.getUi().ButtonSet.OK);
    throw e;
  }
}

/**
 * Load: runs the filter (E=SEND EMAIL, not logged today) for both Math and Reading dashboards
 * and writes the result as values into each sheet's I:L. Works from any sheet.
 */
function loadToWorkArea() {
  try {
    debugLog('Load', 'loadToWorkArea start');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Log');
    if (!logSheet) {
      SpreadsheetApp.getUi().alert('No Log sheet', 'Create a Log sheet first (run Move once to create it).', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const todayCell = logSheet.getRange(1, 20);
    todayCell.setFormula('=TODAY()');
    SpreadsheetApp.flush();
    const today = todayCell.getValue();
    todayCell.clearContent();

    const todayT = (today && today.getTime) ? today.getTime() : (typeof today === 'number' ? today : 0);
    const todayDay = Math.floor(todayT / 86400000);
    debugLog('Load', 'today', { todayDay: todayDay });

    const logLastRow = Math.max(logSheet.getLastRow(), 1);
    const logNumRows = Math.max(logLastRow - 1, 0);
    const logData = logNumRows > 0 ? logSheet.getRange(2, 9, logNumRows, 6).getValues() : []; // I:N = 6 cols (getRange row, col, numRows, numCols)
    if (logData.length > 0 && logData[0].length !== 6) {
      debugLog('Load', 'Log I:N column count', { rows: logData.length, cols: logData[0].length, expected: 6 });
    }
    const loggedTodayBySubject = { Math: {}, Reading: {} };
    const issueRowsBySubject = { Math: [], Reading: [] };
    const allIssueLoginIdsBySubject = { Math: {}, Reading: {} };
    for (let r = 0; r < logData.length; r++) {
      const row = logData[r];
      const subj = String(row[0] || '').trim();
      if (subj !== 'Math' && subj !== 'Reading') continue;
      const loginId = String(row[1] || '');
      allIssueLoginIdsBySubject[subj][loginId] = true;
      issueRowsBySubject[subj].push({
        loginId: loginId,
        name: String(row[2] || ''),
        triggerNum: row[3] != null ? row[3] : '',
        note: String(row[4] || '')
      });
      const d = row[5];
      if (d != null) {
        const t = (d && d.getTime) ? d.getTime() : (typeof d === 'number' ? d : 0);
        if (Math.floor(t / 86400000) === todayDay) {
          loggedTodayBySubject[subj][loginId] = true;
        }
      }
    }
    debugLog('Load', 'Log parsed', { mathIssues: issueRowsBySubject.Math.length, readingIssues: issueRowsBySubject.Reading.length });

    let mathCount = 0;
    let readingCount = 0;

    const mathSheet = ss.getSheetByName('Math Dashboard') || findSheetByName(ss, 'math', 'dashboard');
    if (mathSheet) {
      mathCount = loadOneDashboard(mathSheet, loggedTodayBySubject.Math || {}, issueRowsBySubject.Math || [], allIssueLoginIdsBySubject.Math || {}, 500);
    }

    const readingSheet = ss.getSheetByName('Reading Dashboard') || findSheetByName(ss, 'reading', 'dashboard');
    if (readingSheet) {
      readingCount = loadOneDashboard(readingSheet, loggedTodayBySubject.Reading || {}, issueRowsBySubject.Reading || [], allIssueLoginIdsBySubject.Reading || {}, 500);
    }

    SpreadsheetApp.getUi().alert(
      'Load complete',
      'Math: ' + mathCount + ' rows\nReading: ' + readingCount + ' rows\nIssue rows from the Log are restored with Status=Issue and Notes. Edit and send to Sent or Move back to Log.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    debugLog('Load', 'ERROR', { message: e.message, stack: e.stack });
    SpreadsheetApp.getUi().alert('Load error', e.message + '\n\nCheck View > Logs (Execution log) for details.', SpreadsheetApp.getUi().ButtonSet.OK);
    throw e;
  }
}

function findSheetByName(ss, word1, word2) {
  const w1 = word1.toLowerCase();
  const w2 = word2.toLowerCase();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const n = sheets[i].getName().toLowerCase();
    if (n.indexOf(w1) !== -1 && n.indexOf(w2) !== -1) return sheets[i];
  }
  return null;
}

// Work area I:N = 6 columns (LoginID, Name, col3, Trigger#, Status, Notes)
const WORK_AREA_COLS = 6;
const WORK_AREA_START_ROW = 3;
const WORK_AREA_START_COL = 9; // I

function debugLog(context, message, detail) {
  const d = detail != null ? ' | ' + JSON.stringify(detail) : '';
  Logger.log('[KNA ' + context + '] ' + message + d);
}

function loadOneDashboard(sheet, loggedTodayIds, issueRowsFromLog, allIssueLoginIds, clearMaxRows) {
  const sheetName = sheet.getName();
  debugLog('Load', 'loadOneDashboard start', { sheet: sheetName });
  const lastRow = sheet.getLastRow();
  const numExistingRows = lastRow >= WORK_AREA_START_ROW ? Math.min(lastRow - WORK_AREA_START_ROW + 1, clearMaxRows) : 0;
  const existingRange = numExistingRows > 0
    ? sheet.getRange(WORK_AREA_START_ROW, WORK_AREA_START_COL, numExistingRows, WORK_AREA_COLS).getValues()
    : [];
  debugLog('Load', 'existing work area read', { sheet: sheetName, rows: existingRange.length, cols: existingRange[0] ? existingRange[0].length : 0 });
  const leftover = [];
  const existingIds = {};
  for (let r = 0; r < existingRange.length; r++) {
    const row = existingRange[r];
    const id = String(row[0] || '').trim();
    if (!id) continue;
    leftover.push(row.slice(0, WORK_AREA_COLS));
    existingIds[id] = true;
  }
  const data = lastRow >= 3 ? sheet.getRange(3, 1, lastRow, 7).getValues() : [];
  const out = [];
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (String((row[4] || '')).trim().toLowerCase() !== 'send email') continue;
    const id = String(row[0] || '');
    if (loggedTodayIds[id]) continue;
    if (allIssueLoginIds && allIssueLoginIds[id]) continue; // don't show in SEND EMAIL if they're in the Issue log (any date)
    out.push([row[0], row[1], row[5], row[6], 'Not Sent', '']);
  }
  const merged = leftover.slice();
  const addedIds = {};
  for (let i = 0; i < leftover.length; i++) {
    addedIds[String(leftover[i][0] || '')] = true;
  }
  for (let i = 0; i < (issueRowsFromLog || []).length; i++) {
    const x = issueRowsFromLog[i];
    if (!x || !x.loginId || addedIds[x.loginId]) continue;
    addedIds[x.loginId] = true;
    merged.push([x.loginId, x.name, '', x.triggerNum, 'Issue', x.note || '']);
  }
  for (let i = 0; i < out.length; i++) {
    const row = out[i];
    if (addedIds[row[0]]) continue;
    merged.push(row);
  }
  debugLog('Load', 'merge counts', { sheet: sheetName, leftover: leftover.length, issue: (issueRowsFromLog || []).length, fromDashboard: out.length, merged: merged.length });
  if (merged.length > 0) {
    const cols = merged[0].length;
    if (cols !== WORK_AREA_COLS) {
      debugLog('Load', 'ERROR: merged row column count mismatch', { sheet: sheetName, mergedRows: merged.length, mergedCols: cols, expected: WORK_AREA_COLS });
      throw new Error('Load [' + sheetName + ']: data has ' + cols + ' columns but work area expects ' + WORK_AREA_COLS + '. merged.length=' + merged.length + '. Check leftover and new rows have exactly 6 cells each.');
    }
    const endRow = Math.max(lastRow, 2 + clearMaxRows);
    sheet.getRange(WORK_AREA_START_ROW, WORK_AREA_START_COL, endRow, WORK_AREA_COLS).clearContent();
    sheet.getRange(WORK_AREA_START_ROW, WORK_AREA_START_COL, merged.length, WORK_AREA_COLS).setValues(merged);
    debugLog('Load', 'write done', { sheet: sheetName, rows: merged.length, cols: WORK_AREA_COLS });
  }
  return merged.length;
}

/**
 * Finds 1-based column indices for LoginID, Name, Trigger #, Email, Status, Notes.
 * @param {any[]} headerRow - Row 1 values
 * @returns {{ loginId: number, name: number, triggerNum: number, email: number, status: number, notes: number }}
 */
function getColumnIndices(headerRow) {
  const out = { loginId: 0, name: 0, triggerNum: 0, email: 0, status: 0, notes: 0 };
  for (let c = 0; c < headerRow.length; c++) {
    const h = String(headerRow[c] || '').trim().toLowerCase();
    if (h === 'loginid') out.loginId = c + 1;
    else if (h === 'name') out.name = c + 1;
    else if (h === 'trigger #' || h.indexOf('trigger') !== -1) out.triggerNum = c + 1;
    else if (h === 'email') out.email = c + 1;
    else if (h === 'status') out.status = c + 1;
    else if (h === 'notes') out.notes = c + 1;
  }
  if (!out.triggerNum) {
    for (let c = 0; c < headerRow.length; c++) {
      if (String(headerRow[c]).indexOf('Trigger') !== -1) { out.triggerNum = c + 1; break; }
    }
  }
  return out;
}

/**
 * Returns the last row with data in the given column (1-based). 0 if column is empty.
 */
function getNextLogRow(logSheet, colA1) {
  const col = logSheet.getRange(1, colA1, logSheet.getMaxRows(), colA1);
  const vals = col.getValues();
  let last = 0;
  for (let r = vals.length - 1; r >= 0; r--) {
    if (vals[r][0] !== null && String(vals[r][0]).trim() !== '') {
      last = r + 1;
      break;
    }
  }
  return last;
}
