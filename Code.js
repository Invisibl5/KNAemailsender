/**
 * KNAemailsender - Kumon Email Sender
 *
 * Auto-synced from GitHub. Import folder: CSVs dropped into a Drive folder
 * are imported to "Math Data" and "Reading Data" tabs, then archived.
 */

// --- Import folder config ---
const IMPORT_FOLDER_NAME = 'KNA Email Sender Import';
const ARCHIVE_FOLDER_NAME = 'Archive';

/**
 * Runs when the spreadsheet is opened. Adds a custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KNA Email Sender')
    .addItem('Import from Drive', 'importFromDrive')
    .addItem('Create / Open Import Folder', 'createOrOpenImportFolder')
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
 * Imports a Drive CSV file into a sheet: clears the sheet and writes all rows.
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
