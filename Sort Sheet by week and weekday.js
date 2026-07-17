function createOldExtensionMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Task Tools')
    .addItem('Sort by Week & Day', 'sortTasks')
    .addToUi();
}

function sortTasks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const HEADER_ROWS = 3;   // rows 1–3 are headers
  const WEEK_COL = 6;      // column F
  const WEEKDAY_COL = 7;   // column G
  const lastRow = sheet.getRange("A4").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  if (lastRow <= HEADER_ROWS) return;

  const lastCol = 9;       // columns A–I
  const helperCol = lastCol + 1;

  // Map weekday names
  const dayMap = {
    'monday':1,'mon':1,
    'tuesday':2,'tue':2,'tues':2,
    'wednesday':3,'wed':3,
    'thursday':4,'thu':4,'thurs':4,
    'friday':5,'fri':5,
    'saturday':6,'sat':6,
    'sunday':7,'sun':7,
    'måndag':1,'tisdag':2,'onsdag':3,'torsdag':4,'fredag':5,'lördag':6,'söndag':7,
    'lordag':6,'sondag':7
  };

  // Build weekday numbers
  const weekdayRange = sheet.getRange(HEADER_ROWS + 1, WEEKDAY_COL, lastRow - HEADER_ROWS, 1);
  const weekdayValues = weekdayRange.getValues();
  const helperValues = weekdayValues.map(r => {
    const v = r[0];
    if (v === null || v === '') return [999];
    if (typeof v === 'number') {
      const n = Math.floor(v);
      if (n >= 1 && n <= 7) return [n];
    }
    const s = String(v).trim().toLowerCase();
    return [ dayMap[s] || 999 ];
  });

  // Write helper column
  sheet.getRange(HEADER_ROWS + 1, helperCol, helperValues.length, 1).setValues(helperValues);

  // Sort rows 4..lastRow
  sheet.getRange(HEADER_ROWS + 1, 1, lastRow - HEADER_ROWS, helperCol)
       .sort([{column: WEEK_COL, ascending: true}, {column: helperCol, ascending: true}]);

  // Remove helper column
  sheet.deleteColumn(helperCol);

  // --- Add black borders at week boundaries ---
  const weeks = sheet.getRange(HEADER_ROWS + 1, WEEK_COL, lastRow - HEADER_ROWS, 1).getValues();
  let prevWeek = null;
  for (let i = 0; i < weeks.length; i++) {
    const row = HEADER_ROWS + 1 + i;
    const currentWeek = weeks[i][0];
    if (currentWeek !== prevWeek) {
      // Apply thick black top border to full row A–H
      sheet.getRange(row, 1, 1, lastCol).setBorder(true, null, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
    } else {
      // Clear top border if not a new week
      sheet.getRange(row, 1, 1, lastCol).setBorder(false, null, null, null, null, null);
    }
    prevWeek = currentWeek;
  }
}
