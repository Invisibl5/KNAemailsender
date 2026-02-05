/**
 * KNAemailsender - Kumon Email Sender
 *
 * Auto-synced from GitHub. Import folder: CSVs dropped into a Drive folder
 * are imported to "Math Data" and "Reading Data" tabs, then archived.
 */

// --- Version (bump when you deploy changes) ---
const VERSION = '1.0.9';

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
    .addItem('Sync Dashboard to Log', 'syncDashboardToLog')
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
// Log button only copies to the Log (does not clear Status/Notes). Use your FILTER formula to hide logged rows.

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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const name = sheet.getName();

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

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = getColumnIndices(headerRow);
  // Use the LoginID/Name/Trigger # block that contains Status (second block: I,J,K = 9,10,11 when Status in M)
  if (col.status >= 4) {
    col.loginId = col.status - 4;
    col.name = col.status - 3;
    col.triggerNum = col.status - 2;
  }

  if (!col.loginId || !col.status) {
    SpreadsheetApp.getUi().alert(
      'Missing headers',
      'Dashboard must have headers: LoginID, Name, Trigger #, Email, Status, Notes (in row 1).',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const issueRows = [];
  const issueSheetRows = []; // 1-based sheet row numbers for rows we log (clear Status/Notes after)
  const sentMathRows = [];
  const sentReadingRows = [];

  const dataRange = sheet.getRange(2, 1, lastRow, sheet.getLastColumn());
  const rows = dataRange.getValues();

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const status = String((row[col.status - 1] || '')).trim();

    if (status === '' || status.toLowerCase() === 'not sent') continue;

    const loginId = row[col.loginId - 1];
    const studentName = row[col.name - 1];
    const triggerNum = row[col.triggerNum - 1];
    const note = (col.notes && row[col.notes - 1] != null) ? row[col.notes - 1] : '';

    if (status.toLowerCase() === 'issue') {
      issueRows.push([subject, loginId, studentName, triggerNum, note]);
      issueSheetRows.push(2 + i);
    } else if (status.toLowerCase() === 'sent') {
      if (isMath) {
        sentMathRows.push([loginId, studentName, triggerNum]);
      } else {
        sentReadingRows.push([loginId, studentName, triggerNum]);
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
    const today = new Date();
    today.setHours(0, 0, 0, 0);
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

  const msg = [
    subject + ' Dashboard → Log',
    'Sent: ' + (isMath ? sentMathRows.length : sentReadingRows.length),
    'Issues: ' + issueRows.length
  ].join('\n');
  SpreadsheetApp.getUi().alert('Sync complete', msg, SpreadsheetApp.getUi().ButtonSet.OK);
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
