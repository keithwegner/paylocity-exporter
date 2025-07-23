const EXPORT_FOLDER_ID = '1cCYP1UftlfEd-YlAiL7Gg9uD2geChGPT'; // <- set once
const EMPLOYEE_REGISTRY_ID = '1HAHc8GsxzmsSYtuC42OGyHFMGKxlU6MjpI3yTRFBfnw';
const TIMESHEET_LIST_TAB = 'Current';
const LIST_COL_HEADER = 'Current Timesheet';  // exact header text

/** Main method. 
 *  Executes script across all timesheets found in the employee-registry,
 *  and dumps output file into the EXPORT_FOLDER_ID
 */
function exportAllTimesheets() {

  // Fetch the list of timesheet URLs / IDs
  const links = getTimesheetLinks();

  // Cache of already‑opened Paylocity export files this run
  const exportCache = new Map(); // key = fileName, value = {ss, sheet}

  // Process each timesheet
  links.forEach(link => {
    try {
      const rowsObj = extractPriorPayPeriodRows(link.timesheetId, link.personName);
      if (rowsObj.rows.length === 0) return; // nothing to write

      const target = getPaylocitySheet(rowsObj.periodEndDate, exportCache);
      appendRows(target.sheet, rowsObj.rows);

    } catch (err) {
      Logger.log('✖ Error processing %s -> %s', link.timesheetId, err);
    }
  });
}

/* --------------------------------------------------------------------------
   HELPER FUNCTIONS
   -------------------------------------------------------------------------- */

/**
 * Read master sheet -> return [{timesheetId, personName}, …]
 */
function getTimesheetLinks() {
  const listSS = SpreadsheetApp.openById(EMPLOYEE_REGISTRY_ID);
  const sheet = listSS.getSheetByName(TIMESHEET_LIST_TAB);

  const headerRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];  // row 2
  const richRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getRichTextValues()[0];

  // find “Current Timesheet” column (AN)
  const colIdx = headerRow.findIndex(h =>
    String(h).trim().toUpperCase() === LIST_COL_HEADER.toUpperCase());
  if (colIdx === -1)
    throw new Error(`Column "${LIST_COL_HEADER}" not found in row 2.`);

  // get all values from row 4 downward in that column
  const numRows = sheet.getLastRow() - 3; // rows 4…end
  if (numRows <= 0) return [];

  const colRange = sheet.getRange(4, colIdx + 1, numRows, 1);
  const colValues = colRange.getValues();  // plain text
  const colRichTxt = colRange.getRichTextValues(); // capture embedded links

  const links = [];
  for (let i = 0; i < numRows; i++) {
    const rich = colRichTxt[i][0];
    const url = rich?.getLinkUrl() || String(colValues[i][0]).trim();
    if (!url) continue;

    const idMatch = url.match(/[-\w]{25,}/); // extract spreadsheet ID
    if (idMatch) links.push({ timesheetId: idMatch[0] });
  }
  return links;
}

/**
 * Open a timesheet → return {periodEndDate:Date, rows:Array<Array>}
 * Rows exclude header; ready to append.
 */
function extractPriorPayPeriodRows(timesheetId) {
  const ss = SpreadsheetApp.openById(timesheetId);
  const sheet = ss.getSheets()[0];

  // get both raw and display so we keep leading zeros
  const range = sheet.getDataRange();
  const data = range.getValues();
  const display = range.getDisplayValues();

  /* ---- locate project columns ---- */
  const PROJ_NUM_ROW = 0, TASK_NUM_ROW = 1, WORK_TYPE_ROW = 2, RES_NUM_ROW = 3;
  const skipLabels = ['Total', 'Holiday', 'Overhead', 'Reviewer'];
  const projectCols = [];

  for (let c = 1; c < data[0].length; c++) {
    const projNumDisp = display[PROJ_NUM_ROW][c];
    const workType = String(display[WORK_TYPE_ROW][c]).trim();

    if (skipLabels.includes(String(display[4][c]).trim())) continue;

    if (projNumDisp || ['PTO (used)', 'PTU (used)', 'MBE'].includes(workType)) {
      projectCols.push({
        col: c,
        projNumText: projNumDisp || '', // always text (to keeps 0‑padding)
        taskNum: display[TASK_NUM_ROW][c],
        workType,
        personnelCode: display[RES_NUM_ROW][c],
      });
    }
  }
  if (!projectCols.length) return { periodEndDate: null, rows: [] };

  /* ---- completed END PERIOD range ---- */
  const today = new Date();
  const endRows = [];
  data.forEach((row, idx) => {
    if (String(row[0]).toUpperCase().trim() === 'END PERIOD') {
      const prevDate = data[idx - 1]?.[0];
      if (prevDate instanceof Date && prevDate <= today) endRows.push({ idx, prevDate });
    }
  });
  if (endRows.length < 2) return { periodEndDate: null, rows: [] };

  const lastEnd = endRows[endRows.length - 1].idx;
  const prevEnd = endRows[endRows.length - 2].idx;
  const firstData = prevEnd + 1;
  const lastData = lastEnd - 1;
  const periodEndDate = data[lastData][0];

  /* ---- build export rows ---- */
  const [last, first] = ss.getName().split('_')[0].split('-');
  const personName = `${last}, ${first}`;

  const rows = [];
  for (let r = firstData; r <= lastData; r++) {
    const dateVal = data[r][0];
    if (!(dateVal instanceof Date)) continue;

    projectCols.forEach(p => {
      // keep minus sign, decimal; remove other junk
      const numStr = String(data[r][p.col]).replace(/[^0-9.\-]/g, '');
      const hrs = parseFloat(numStr);
      if (!isNaN(hrs) && hrs > 0) {// skip zero or negative PTO/MBE
        rows.push([
          dateVal,
          p.personnelCode,
          personName,
          hrs,
          "'" + p.projNumText, // add a leading apostrophe to force text output
          p.taskNum,
          p.workType,
        ]);
      }
    });
  }
  return { periodEndDate, rows };
}


/**
 * Find or create the Paylocity_YYYYMMDD file in the export folder.
 * Keeps a cache to avoid reopening when multiple employees share the same period.
 */
function getPaylocitySheet(periodEndDate, cache) {
  const fileName = 'Paylocity_' + Utilities.formatDate(
    periodEndDate, Session.getScriptTimeZone(), 'yyyyMMdd');

  if (cache.has(fileName)) return cache.get(fileName);

  const folder = DriveApp.getFolderById(EXPORT_FOLDER_ID);
  const files = folder.getFilesByName(fileName);
  let ss, isNew = false;

  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(fileName);
    folder.addFile(DriveApp.getFileById(ss.getId()));          // move to folder
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(ss.getId()));
    isNew = true;
  }
  const sheet = ss.getSheets()[0];
  if (isNew && sheet.getLastRow() === 0) {
    sheet.appendRow(['Work Date', 'Personnel Code', 'Person name', 'Hours',
      'Project No.', 'Project Task No.', 'Work Type']);
    sheet.getRange('A2').setNumberFormat('MM/dd/yyyy');
  }

  const info = { ss, sheet };
  cache.set(fileName, info);
  return info;
}

/**
 * Append rows to the Paylocity sheet; keeps date column formatted.
 */
function appendRows(sheet, rows) {
  if (rows.length === 0) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(`A${startRow}:A${startRow + rows.length - 1}`)
    .setNumberFormat('MM/dd/yyyy');
}
