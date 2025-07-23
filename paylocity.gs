function exportPriorPayPeriod () {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const data  = sheet.getDataRange().getValues();          // entire sheet

  // ───── 1. Identify project columns ────────────────────────────────────────────
  const PROJ_NUM_ROW = 0, TASK_NUM_ROW = 1, WORK_TYPE_ROW = 2, RES_NUM_ROW = 3;
  const projectCols  = [];

  for (let c = 1; c < data[PROJ_NUM_ROW].length; c++) {
    const projNum = data[PROJ_NUM_ROW][c];
    if (!projNum) break;                                   // stop at first blank cell

    // (optional) skip “totals” columns
    if (['Total', 'Holiday', 'Overhead', 'Reviewer'].includes(String(data[4][c])))
      continue;

    projectCols.push({
      col: c,
      projNum,
      taskNum:      data[TASK_NUM_ROW][c],
      workType:     data[WORK_TYPE_ROW][c],
      personnelCode:data[RES_NUM_ROW][c],
    });
  }

  Logger.log('🛈 Project columns found: %s', JSON.stringify(projectCols, null, 2));

// ───── 2. Locate the two most‑recent END PERIOD rows *up to today* ────────────
const today = new Date();

// Gather every END PERIOD plus the date that precedes it (the period’s last work‑day)
const endRows = [];
data.forEach((row, idx) => {
  if (String(row[0]).toUpperCase().trim() === 'END PERIOD') {
    const prevDate = data[idx - 1]?.[0];            // date on the row just above
    if (prevDate instanceof Date && prevDate <= today) {
      endRows.push({ idx, prevDate });              // keep only “past” periods
    }
  }
});

Logger.log('🛈 Eligible END PERIOD rows (≤ today): %s',
           endRows.map(r => r.idx + 1));            // 1‑based for display

if (endRows.length < 2)
  throw new Error('Need at least two completed pay periods before today.');

const lastEnd   = endRows[endRows.length - 1].idx;  // zero‑based index
const prevEnd   = endRows[endRows.length - 2].idx;
const firstData = prevEnd + 1;
const lastData  = lastEnd - 1;

Logger.log('🛈 Prior (completed) period = sheet rows %s–%s',
           firstData + 1, lastData + 1);            // 1‑based in log

  // ───── 3. Build the output array ──────────────────────────────────────────────
  const [last, first] = ss.getName().split('_')[0].split('-');
  const personName    = `${last}, ${first}`;

  const out = [['Work Date','Personnel Code','Person name','Hours',
                'Project No.','Project Task No.','Work Type']];

  for (let r = firstData; r <= lastData; r++) {
    const dateVal = data[r][0];
    if (!(dateVal instanceof Date)) {
      Logger.log('‑‑ Skipping row %s; A value = %s (type %s)',
                 r + 1, dateVal, typeof dateVal);
      continue;
    }

    projectCols.forEach(p => {
      const raw = data[r][p.col];
      const hrs = raw === '' ? 0 : Number(raw);

      if (hrs > 0) {
        Logger.log('✔ Row %s | %s | Col %s → %s hrs',
                   r + 1, Utilities.formatDate(dateVal, ss.getSpreadsheetTimeZone(), 'MM/dd'),
                   p.col + 1, hrs);

        out.push([
          dateVal,
          p.personnelCode,
          personName,
          hrs,
          p.projNum,
          p.taskNum,
          p.workType
        ]);
      }
    });
  }

  Logger.log('🛈 Total export rows (excluding header): %s', out.length - 1);

  if (out.length === 1) throw new Error('No billable hours found in the prior pay period.');

  // ───── 4. Create the new sheet & write data ───────────────────────────────────
  const tz    = ss.getSpreadsheetTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmm');
  const newSS = SpreadsheetApp.create(`Palocity Export – ${personName} – ${stamp}`);
  const tgt   = newSS.getActiveSheet();

  tgt.getRange(1,1,out.length,out[0].length).setValues(out);
  tgt.getRange('A2:A').setNumberFormat('MM/dd/yyyy');

  try {
    const parent = DriveApp.getFileById(ss.getId()).getParents().next();
    DriveApp.getFileById(newSS.getId()).moveTo(parent);
  } catch (e) {/* ignore if no parent folder */}
}
