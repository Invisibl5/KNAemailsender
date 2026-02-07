/**
 * KNAemailsender - Kumon Email Sender
 *
 * Auto-synced from GitHub. Import folder: CSVs dropped into a Drive folder
 * are imported to "Math Data" and "Reading Data" tabs, then archived.
 */

// --- Version (bump when you deploy changes) ---
const VERSION = '1.0.38';

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
    .addSeparator()
    .addItem('Verify from ClassNavi', 'verifyFromClassNavi')
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

// Log Issue columns: I=Subject, J=LoginID, K=Name, L=Trigger #, M=Note, N=Date
const LOG_ISSUE_START_COL = 9;
const LOG_ISSUE_NUM_COLS = 6;

/**
 * Looks up Email for a student by LoginID from Math Data or Reading Data sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} subject - "Math" or "Reading"
 * @param {string} loginId
 * @returns {string}
 */
function getEmailFromDataSheet(ss, subject, loginId) {
  const sheetName = subject === 'Math' ? 'Math Data' : 'Reading Data';
  const sheet = ss.getSheetByName(sheetName) || findSheetByName(ss, subject.toLowerCase(), 'data');
  if (!sheet || sheet.getLastRow() < 2) return '';
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let col = getColumnIndices(headerRow);
  if (!col.loginId) {
    for (let c = 0; c < headerRow.length; c++) {
      const h = String(headerRow[c] || '').trim().toLowerCase().replace(/\s+/g, '');
      if (h === 'loginid') { col.loginId = c + 1; break; }
    }
  }
  if (!col.email) {
    for (let c = 0; c < headerRow.length; c++) {
      const h = String(headerRow[c] || '').trim().toLowerCase();
      if (h === 'email') { col.email = c + 1; break; }
    }
  }
  if (!col.loginId || !col.email) return '';
  const data = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const idStr = String(loginId).trim();
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (String(row[col.loginId - 1] != null ? row[col.loginId - 1] : '').trim() === idStr) {
      return row[col.email - 1] != null ? String(row[col.email - 1]) : '';
    }
  }
  return '';
}

/**
 * Load: runs the filter (E=SEND EMAIL, not logged today) for both Math and Reading dashboards,
 * PLUS brings back Issue log entries (Status=Issue, Note) with email lookup.
 * Move sends to log; Load brings from log. Removes brought-back entries from Issue log and compacts.
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

    const logLastRow = Math.max(logSheet.getLastRow(), 1);
    const logNumRows = Math.max(logLastRow - 1, 0);
    const logData = logNumRows > 0 ? logSheet.getRange(2, LOG_ISSUE_START_COL, logNumRows, LOG_ISSUE_NUM_COLS).getValues() : [];
    if (logData.length > 0 && logData[0].length !== LOG_ISSUE_NUM_COLS) {
      debugLog('Load', 'Log I:N column count', { rows: logData.length, cols: logData[0].length, expected: LOG_ISSUE_NUM_COLS });
    }

    // Parse Issue log: Subject, LoginID, Name, Trigger #, Note, Date (I-N)
    // issueNoteByLoginIdAndTrigger[subject][loginId][triggerNum] = note (for dashboard rows that are in log)
    // issueEntriesBySubject[subject] = [{loginId, name, triggerNum, note, sheetRow}]
    const issueNoteByLoginIdAndTrigger = { Math: {}, Reading: {} };
    const issueEntriesBySubject = { Math: [], Reading: [] };
    for (let r = 0; r < logData.length; r++) {
      const row = logData[r];
      const subjRaw = String(row[0] || '').trim();
      const subj = subjRaw.toLowerCase() === 'math' ? 'Math' : (subjRaw.toLowerCase() === 'reading' ? 'Reading' : null);
      if (!subj) continue;
      const loginId = String(row[1] != null ? row[1] : '').trim();
      if (!loginId) continue;
      const triggerNum = row[3];
      const note = String(row[4] || '');
      const sheetRow = 2 + r;
      if (!issueNoteByLoginIdAndTrigger[subj][loginId]) {
        issueNoteByLoginIdAndTrigger[subj][loginId] = {};
      }
      issueNoteByLoginIdAndTrigger[subj][loginId][normalizeTrigger(triggerNum)] = note;
      issueEntriesBySubject[subj].push({ loginId: loginId, name: row[2], triggerNum: triggerNum, note: note, sheetRow: sheetRow });
    }
    debugLog('Load', 'Log parsed', {
      mathIssues: issueEntriesBySubject.Math.length,
      readingIssues: issueEntriesBySubject.Reading.length
    });

    // Don't exclude logged-today: Move sends to log, Load brings from log (including today)
    const loggedTodayIds = { Math: {}, Reading: {} };

    const broughtBackKeys = []; // {subject, loginId, triggerNum} for compacting Log

    let mathCount = 0;
    let readingCount = 0;

    const mathSheet = ss.getSheetByName('Math Dashboard') || findSheetByName(ss, 'math', 'dashboard');
    if (mathSheet) {
      mathCount = loadOneDashboard(mathSheet, loggedTodayIds.Math, issueNoteByLoginIdAndTrigger.Math, issueEntriesBySubject.Math, ss, 'Math', broughtBackKeys, 500);
    }

    const readingSheet = ss.getSheetByName('Reading Dashboard') || findSheetByName(ss, 'reading', 'dashboard');
    if (readingSheet) {
      readingCount = loadOneDashboard(readingSheet, loggedTodayIds.Reading, issueNoteByLoginIdAndTrigger.Reading, issueEntriesBySubject.Reading, ss, 'Reading', broughtBackKeys, 500);
    }

    if (broughtBackKeys.length > 0) {
      compactIssueLog(logSheet, broughtBackKeys, logData);
    }

    SpreadsheetApp.getUi().alert(
      'Load complete',
      'Math: ' + mathCount + ' rows\nReading: ' + readingCount + ' rows\nFilter: SEND EMAIL + Issue log (brought back). Brought ' + broughtBackKeys.length + ' from Issue log.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    debugLog('Load', 'ERROR', { message: e.message, stack: e.stack });
    SpreadsheetApp.getUi().alert('Load error', e.message + '\n\nCheck View > Logs (Execution log) for details.', SpreadsheetApp.getUi().ButtonSet.OK);
    throw e;
  }
}

/** Normalize trigger for comparison (string or number). */
function normalizeTrigger(t) {
  if (t == null) return '';
  return String(t).trim();
}

/**
 * Removes Issue log entries that were brought back to dashboards, then compacts I-N by shifting rows up.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} logSheet
 * @param {Array<{subject: string, loginId: string, triggerNum: *}>} keysToRemove
 * @param {Array<Array>} logData - full I:N data (from row 2)
 */
function compactIssueLog(logSheet, keysToRemove, logData) {
  const removeSet = {};
  for (let i = 0; i < keysToRemove.length; i++) {
    const k = keysToRemove[i];
    const key = k.subject + '|' + String(k.loginId).trim() + '|' + normalizeTrigger(k.triggerNum);
    removeSet[key] = true;
  }
  const remaining = [];
  for (let r = 0; r < logData.length; r++) {
    const row = logData[r];
    const subjRaw = String(row[0] || '').trim();
    const subj = subjRaw.toLowerCase() === 'math' ? 'Math' : (subjRaw.toLowerCase() === 'reading' ? 'Reading' : null);
    if (!subj) continue;
    const loginId = String(row[1] != null ? row[1] : '').trim();
    if (!loginId) continue;
    const triggerNum = row[3];
    const key = subj + '|' + loginId + '|' + normalizeTrigger(triggerNum);
    if (removeSet[key]) continue;
    remaining.push(row);
  }
  const firstRow = 2;
  const lastRow = firstRow + logData.length - 1;
  if (lastRow < firstRow) return;
  logSheet.getRange(firstRow, LOG_ISSUE_START_COL, lastRow, LOG_ISSUE_START_COL + LOG_ISSUE_NUM_COLS - 1).clearContent();
  if (remaining.length > 0) {
    logSheet.getRange(firstRow, LOG_ISSUE_START_COL, remaining.length, LOG_ISSUE_NUM_COLS).setValues(remaining);
  }
  debugLog('Load', 'compactIssueLog', { removed: keysToRemove.length, remaining: remaining.length });
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

function loadOneDashboard(sheet, loggedTodayIds, issueNoteByLoginIdAndTrigger, issueEntriesFromLog, ss, subject, broughtBackKeys, clearMaxRows) {
  const sheetName = sheet.getName();
  debugLog('Load', 'loadOneDashboard start', { sheet: sheetName });
  const lastRow = sheet.getLastRow();
  const numExistingRows = lastRow >= WORK_AREA_START_ROW ? Math.min(lastRow - WORK_AREA_START_ROW + 1, clearMaxRows) : 0;
  const existingRange = numExistingRows > 0
    ? sheet.getRange(WORK_AREA_START_ROW, WORK_AREA_START_COL, numExistingRows, WORK_AREA_COLS).getValues()
    : [];
  debugLog('Load', 'existing work area read', { sheet: sheetName, rows: existingRange.length, cols: existingRange[0] ? existingRange[0].length : 0 });
  const leftover = [];
  const addedKeys = {};
  function key(loginId, triggerNum) {
    return String(loginId || '').trim() + '|' + normalizeTrigger(triggerNum);
  }
  for (let r = 0; r < existingRange.length; r++) {
    const row = existingRange[r];
    const id = String(row[0] || '').trim();
    if (!id) continue;
    const tr = row[3];
    leftover.push(row.slice(0, WORK_AREA_COLS));
    addedKeys[key(id, tr)] = true;
  }
  const data = lastRow >= 3 ? sheet.getRange(3, 1, lastRow, 7).getValues() : [];
  const headerRow = lastRow >= 2 ? sheet.getRange(2, 1, 2, Math.max(sheet.getLastColumn(), 7)).getValues()[0] : [];
  const col = getColumnIndices(headerRow);
  const emailCol = (col.email && col.email >= 1 && col.email <= 7) ? col.email - 1 : 5;
  const out = [];
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (String((row[4] || '')).trim().toLowerCase() !== 'send email') continue;
    const id = String(row[0] != null ? row[0] : '').trim();
    if (!id) continue;
    if (loggedTodayIds[id]) continue;
    const triggerNum = row[6];
    const email = row[emailCol] != null ? String(row[emailCol]) : '';
    const noteMap = issueNoteByLoginIdAndTrigger && issueNoteByLoginIdAndTrigger[id];
    const note = noteMap && noteMap[normalizeTrigger(triggerNum)] != null ? String(noteMap[normalizeTrigger(triggerNum)]) : '';
    const status = note ? 'Issue' : 'Not Sent';
    out.push([row[0], row[1], email, triggerNum, status, note]);
  }
  const merged = leftover.slice();
  for (let i = 0; i < out.length; i++) {
    const row = out[i];
    const rowId = String(row[0] != null ? row[0] : '').trim();
    const tr = row[3];
    if (addedKeys[key(rowId, tr)]) continue;
    addedKeys[key(rowId, tr)] = true;
    merged.push(row);
  }
  for (let i = 0; i < issueEntriesFromLog.length; i++) {
    const ent = issueEntriesFromLog[i];
    const id = String(ent.loginId).trim();
    const tr = ent.triggerNum;
    if (addedKeys[key(id, tr)]) continue;
    addedKeys[key(id, tr)] = true;
    const email = getEmailFromDataSheet(ss, subject, id);
    merged.push([ent.loginId, ent.name, email, ent.triggerNum, 'Issue', ent.note]);
    broughtBackKeys.push({ subject: subject, loginId: id, triggerNum: tr });
  }
  debugLog('Load', 'merge counts', { sheet: sheetName, leftover: leftover.length, fromFilter: out.length, fromIssueLog: issueEntriesFromLog.length, merged: merged.length });
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

// --- ClassNavi verify (page verification from Kumon API) ---
const CLASSNAVI_BASE_URL = 'https://instructor2.digital.kumon.com/USA';
const CLASSNAVI_LOGIN_ID = '404653110035';
const CLASSNAVI_PASSWORD_HASH = 't4lBET8jlgtti7RBBY2tOMrZJ%2FvY2Nfe3Qgt4eI7FEI%3D';
const CLASSNAVI_STATUS_CELL_ROW = 1;
const CLASSNAVI_STATUS_CELL_COL = 20; // T
const CLASSNAVI_RESULT_START_COL = 16; // P = Level, Q = Lowest From, R = Lowest To, S = Error
const CLASSNAVI_INACTIVE_MATH_COL = 20;   // T = Math Inactive (Y/N)
const CLASSNAVI_INACTIVE_READING_COL = 21; // U = Reading Inactive (Y/N)

function classNaviClientObject(id) {
  if (id == null) id = Date.now();
  return {
    applicationName: 'Class-Navi',
    version: '1.0.0.0',
    programName: 'Class-Navi',
    machineName: '-',
    os: 'Web',
    id: String(id)
  };
}

function classNaviLogin(loginId, passwordHash) {
  const usernameEncoded = 'USA%2F' + encodeURIComponent(String(loginId).trim());
  // Hash from NaviPasswordHash cookie is already URL-encoded; use as-is
  const body = 'grant_type=password&username=' + usernameEncoded + '&password=' + String(passwordHash).trim();
  const res = UrlFetchApp.fetch(CLASSNAVI_BASE_URL + '/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: body,
    muteHttpExceptions: true,
    headers: {
      'Accept': 'application/json, text/plain, */*',
      'Origin': 'https://instructor2.digital.kumon.com',
      'Referer': 'https://instructor2.digital.kumon.com/USA/'
    }
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code !== 200) {
    throw new Error('ClassNavi login failed (' + code + '): ' + text.substring(0, 200));
  }
  const data = JSON.parse(text);
  if (!data || !data.access_token) {
    throw new Error('ClassNavi login: no access_token. Get a fresh hash from NaviPasswordHash cookie.');
  }
  return data.access_token;
}

function classNaviApiCall(token, endpoint, body) {
  const res = UrlFetchApp.fetch(CLASSNAVI_BASE_URL + endpoint, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
    headers: { 'Authorization': 'Bearer ' + token }
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code !== 200) {
    throw new Error('ClassNavi API failed (' + code + '): ' + text.substring(0, 300));
  }
  const data = JSON.parse(text);
  if (data.Result && data.Result.ResultCode !== 0) {
    const errors = (data.Result.Errors || []).map(function (e) { return e.Message || e.ErrorCode; });
    throw new Error('ClassNavi API error: ' + (errors.join(', ') || 'ResultCode ' + data.Result.ResultCode));
  }
  return data;
}

function classNaviGetInstructorInfo(token, loginId) {
  return classNaviApiCall(token, '/api/ATX0010P/GetInstructorInfo', {
    SystemCountryCD: 'USA',
    LoginID: loginId,
    client: classNaviClientObject()
  });
}

function classNaviExtractStudentList(res) {
  if (!res) return [];
  if (Array.isArray(res)) return res;
  if (res.CenterAllStudentList && Array.isArray(res.CenterAllStudentList)) return res.CenterAllStudentList;
  if (res.StudentInfoList && Array.isArray(res.StudentInfoList)) return res.StudentInfoList;
  if (res.StudentList && Array.isArray(res.StudentList)) return res.StudentList;
  var first = null;
  for (var key in res) {
    if (Array.isArray(res[key])) { first = res[key]; break; }
  }
  return first || [];
}

function classNaviGetAllStudents(token, centerID, instructorID, instructorAssistantSec) {
  var students = [];
  var pageSize = 100;
  var baseBody = {
    SystemCountryCD: 'USA',
    CenterID: centerID,
    InstructorID: instructorID,
    InstructorAssistantSec: instructorAssistantSec,
    ValidFlg: '1',
    client: classNaviClientObject()
  };
  var useStartNum = false;
  var list = classNaviExtractStudentList(classNaviApiCall(token, '/api/ATE0010P/GetCenterAllStudentList', Object.assign({}, baseBody, { Offset: 1, GetNum: pageSize })));
  if (list.length === 0) {
    list = classNaviExtractStudentList(classNaviApiCall(token, '/api/ATE0010P/GetCenterAllStudentList', Object.assign({}, baseBody, { StartNum: 1, DispNum: pageSize })));
    useStartNum = true;
  }
  while (list.length > 0) {
    for (var i = 0; i < list.length; i++) students.push(list[i]);
    if (list.length < pageSize) break;
    Utilities.sleep(500);
    if (useStartNum) {
      list = classNaviExtractStudentList(classNaviApiCall(token, '/api/ATE0010P/GetCenterAllStudentList', Object.assign({}, baseBody, { StartNum: 1 + students.length, DispNum: pageSize })));
    } else {
      list = classNaviExtractStudentList(classNaviApiCall(token, '/api/ATE0010P/GetCenterAllStudentList', Object.assign({}, baseBody, { Offset: 1 + students.length, GetNum: pageSize })));
    }
  }
  return students;
}

function classNaviGetStudyResult(token, studentID, classID, classStudentSeq, subjectCD, centerID, worksheetCD) {
  var body = {
    SystemCountryCD: 'USA',
    StudentID: studentID,
    ClassID: classID,
    ClassStudentSeq: classStudentSeq,
    SubjectCD: subjectCD,
    client: classNaviClientObject()
  };
  if (centerID) body.CenterID = centerID;
  if (worksheetCD) body.WorksheetCD = worksheetCD;
  return classNaviApiCall(token, '/api/ATD0010P/GetStudyResultInfoList', body);
}

function classNaviComputeLowest(data) {
  var list = (data && data.StudyUnitInfoList) ? data.StudyUnitInfoList : [];
  var planned = [];
  for (var i = 0; i < list.length; i++) {
    var u = list[i];
    if (!u || u.StudyStatus === '6') continue;
    if (u.StudyDate || u.FinishDate) continue;
    planned.push(u);
  }
  var minFrom = null, minTo = null;
  for (var i = 0; i < planned.length; i++) {
    var u = planned[i];
    var from = u.WorksheetNOFrom;
    if (from == null || from === '') continue;
    var fromN = Number(from);
    var toN = (u.WorksheetNOTo != null && u.WorksheetNOTo !== '') ? Number(u.WorksheetNOTo) : null;
    if (isNaN(fromN)) continue;
    if (minFrom === null || fromN < minFrom) {
      minFrom = fromN;
      minTo = (toN != null && !isNaN(toN)) ? toN : null;
    }
  }
  return { minFrom: minFrom, minTo: minTo };
}

/**
 * Verify from ClassNavi: combined. For each student in work area (column I), fetches lowest planned page
 * (P–S) and active/inactive status (T) for this dashboard's subject only. Run from Math or Reading Dashboard.
 * Loading shown in cell T1 throughout.
 */
function verifyFromClassNavi() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var name = sheet.getName();
  var isMath = name.toLowerCase().indexOf('math') !== -1;
  var isReading = name.toLowerCase().indexOf('reading') !== -1;
  var isDashboard = name.toLowerCase().indexOf('dashboard') !== -1;
  if (!isDashboard || (!isMath && !isReading)) {
    ui.alert('Wrong sheet', 'Please run "Verify from ClassNavi" from Math Dashboard or Reading Dashboard.', ui.ButtonSet.OK);
    return;
  }
  var subjectCD = isMath ? '010' : '022';
  var colLabel = isMath ? 'Math' : 'Reading';
  var lastRow = sheet.getLastRow();
  if (lastRow < WORK_AREA_START_ROW) {
    ui.alert('No data', 'No rows in work area. Load students first.', ui.ButtonSet.OK);
    return;
  }
  var rowsWithLoginId = [];
  for (var r = WORK_AREA_START_ROW; r <= lastRow; r++) {
    var val = sheet.getRange(r, WORK_AREA_START_COL).getValue();
    var id = (val != null && String(val).trim() !== '') ? String(val).trim() : '';
    if (id) rowsWithLoginId.push({ row: r, loginId: id });
  }
  if (rowsWithLoginId.length === 0) {
    ui.alert('No students', 'No LoginIDs in column I. Load students first.', ui.ButtonSet.OK);
    return;
  }
  var statusRange = sheet.getRange(CLASSNAVI_STATUS_CELL_ROW, CLASSNAVI_STATUS_CELL_COL);
  function setStatus(msg) {
    statusRange.setValue(msg);
    SpreadsheetApp.flush();
  }
  setStatus('Verifying... Logging in... 0/' + rowsWithLoginId.length);
  var token;
  try {
    token = classNaviLogin(CLASSNAVI_LOGIN_ID, CLASSNAVI_PASSWORD_HASH);
  } catch (e) {
    statusRange.clearContent();
    ui.alert('ClassNavi login failed', e.message, ui.ButtonSet.OK);
    return;
  }
  setStatus('Verifying... Fetching student list... 0/' + rowsWithLoginId.length);
  var instructorInfo, centerID, instructorAssistantSec, allStudents;
  try {
    instructorInfo = classNaviGetInstructorInfo(token, CLASSNAVI_LOGIN_ID);
    centerID = instructorInfo.MainCenterID || (instructorInfo.CenterInfoList && instructorInfo.CenterInfoList[0] && instructorInfo.CenterInfoList[0].CenterID);
    instructorAssistantSec = instructorInfo.InstructorAssistantSec || '2';
    allStudents = classNaviGetAllStudents(token, centerID, CLASSNAVI_LOGIN_ID, instructorAssistantSec);
  } catch (e) {
    statusRange.clearContent();
    ui.alert('ClassNavi fetch failed', e.message, ui.ButtonSet.OK);
    return;
  }
  var loginIdToStudent = {};
  for (var i = 0; i < allStudents.length; i++) {
    var s = allStudents[i];
    var lid = (s.LoginID != null ? String(s.LoginID) : '') || (s.StudentID != null ? String(s.StudentID) : '');
    if (lid) loginIdToStudent[lid] = s;
  }
  function getSubjectStatus(studyList, subjCD) {
    if (!studyList || !Array.isArray(studyList)) return 'not enrolled';
    var want = subjCD === '010' || subjCD === 10 ? [10, '010'] : [22, '022'];
    for (var j = 0; j < studyList.length; j++) {
      var s = studyList[j];
      var cd = s.SubjectCD;
      if (cd === want[0] || cd === want[1]) {
        var st = s.Status;
        return (st === '1' || st === 1) ? 'inactive' : 'active';
      }
    }
    return 'not enrolled';
  }
  var headersDone = false;
  for (var i = 0; i < rowsWithLoginId.length; i++) {
    setStatus('Verifying... ' + (i + 1) + '/' + rowsWithLoginId.length);
    var item = rowsWithLoginId[i];
    var r = item.row;
    var loginId = item.loginId;
    var level = '';
    var lowestFrom = '';
    var lowestTo = '';
    var errMsg = '';
    var activeStatus = 'not found';
    var student = loginIdToStudent[loginId];
    if (!student) {
      errMsg = 'Not found in ClassNavi';
    } else {
      var studyList = student.StudentStudyInfoList || [];
      activeStatus = getSubjectStatus(studyList, subjectCD);
      var study = null;
      for (var j = 0; j < studyList.length; j++) {
        if (studyList[j].SubjectCD === subjectCD || studyList[j].SubjectCD === (isMath ? 10 : 22)) { study = studyList[j]; break; }
      }
      if (!study || study.ClassID == null || study.ClassStudentSeq == null) {
        if (!errMsg) errMsg = 'No ' + colLabel + ' study info';
      } else {
        try {
          var result = classNaviGetStudyResult(
            token,
            student.StudentID || student.LoginID,
            study.ClassID,
            study.ClassStudentSeq,
            subjectCD,
            centerID,
            study.NextWorksheetCD
          );
          level = study.NextWorksheetCD != null ? String(study.NextWorksheetCD) : '';
          var lowest = classNaviComputeLowest(result);
          if (lowest.minFrom != null) lowestFrom = lowest.minFrom;
          if (lowest.minTo != null) lowestTo = lowest.minTo;
        } catch (e) {
          errMsg = e.message ? e.message.substring(0, 100) : 'API error';
        }
        Utilities.sleep(400);
      }
    }
    if (!headersDone) {
      sheet.getRange(2, CLASSNAVI_RESULT_START_COL, 1, 5).setValues([['ClassNavi Level', 'Lowest From', 'Lowest To', 'ClassNavi Error', colLabel]]);
      headersDone = true;
    }
    sheet.getRange(r, CLASSNAVI_RESULT_START_COL, 1, 5).setValues([[level, lowestFrom, lowestTo, errMsg, activeStatus]]);
  }
  statusRange.clearContent();
  ui.alert('Verify complete', 'Verified ' + rowsWithLoginId.length + ' students. P–S: Level/Lowest/Error. T: ' + colLabel + ' (active/inactive/not enrolled).', ui.ButtonSet.OK);
}
