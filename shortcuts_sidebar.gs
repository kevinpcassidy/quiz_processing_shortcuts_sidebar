function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Formula Tools")
    .addItem("Open Sidebar", "showSidebar")
    .addToUi();
}

// Show sidebar
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Formula Tools')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ============================
   SERVER FUNCTIONS CALLED BY SIDEBAR
   ============================ */

// 1️⃣ Select columns down to last visible row
function selectColumnsDown() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRangeList();

  if (!selection || selection.getRanges().length < 1) {
    return "Please select at least one cell to start.";
  }

  const ranges = selection.getRanges();
  const newRanges = [];

  ranges.forEach(range => {
    const col = range.getColumn();
    let startRow = range.getRow();
    if (startRow === 1) startRow = 2;

    const filter = sheet.getFilter();
    let lastVisibleRow = sheet.getLastRow();

    if (filter) {
      const maxRow = filter.getFilterRange().getLastRow();
      for (let r = maxRow; r >= startRow; r--) {
        if (!sheet.isRowHiddenByFilter(r)) {
          lastVisibleRow = r;
          break;
        }
      }
    }

    if (startRow <= lastVisibleRow) {
      newRanges.push(sheet.getRange(startRow, col, lastVisibleRow - startRow + 1));
    }
  });

  if (newRanges.length > 0) {
    sheet.getRangeList(newRanges.map(r => r.getA1Notation())).activate();
  }

  return "Selection extended to bottom visible row.";
}

// 2️⃣ Fill average formulas (bulk for speed)
function fillAverageFormulas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveCell();
  const targetCol = range.getColumn();

  if (targetCol <= 4) {
    return "Please select a column at least 5 or later (needs 4 preceding columns).";
  }

  const lastRow = sheet.getLastRow();
  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const startCol = targetCol - 4;

  const template =
    '=IF(COUNTA(%range%)=0,"",IFERROR(AVERAGE(LARGE(%range%,{1}),LARGE(%range%,{2})),MAX(%range%)))';

  const formulas = names.map((row, i) => {
    if (row[0]) {
      const rangeA1 = sheet.getRange(i + 2, startCol, 1, 4).getA1Notation();
      return [template.replace(/%range%/g, rangeA1)];
    } else {
      return [''];
    }
  });

  sheet.getRange(2, targetCol, formulas.length, 1).setFormulas(formulas);
  return "Average formulas filled successfully.";
}

// 3️⃣ Update scores from source sheet
function updateScoresFromSourceSheet(sheetName) {
  if (!sheetName) return "Please select a valid sheet to update.";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sheetName);
  const targetSheet = ss.getActiveSheet();

  if (!sourceSheet) return "Source sheet not found.";

  const sourceLastCol = sourceSheet.getLastColumn();
  const sourceLastRow = sourceSheet.getLastRow();
  const targetLastCol = targetSheet.getLastColumn();
  const targetLastRow = targetSheet.getLastRow();

  if (sourceLastCol < 1 || sourceLastRow < 2)
    return "Source sheet is empty or missing data.";
  if (targetLastCol < 1 || targetLastRow < 2)
    return "Target sheet has no data to update.";

  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceLastCol).getValues()[0];
  const sourceData = sourceSheet.getRange(2, 1, sourceLastRow - 1, sourceLastCol).getValues();
  const sourceNames = sourceData.map(r => r[0]);

  const targetHeaders = targetSheet.getRange(1, 1, 1, targetLastCol).getValues()[0];
  const targetNames = targetSheet.getRange(2, 1, targetLastRow - 1, 1).getValues().map(r => r[0]);

  // Build map of source rows by student name
  const sourceMap = {};
  for (let i = 0; i < sourceNames.length; i++) {
    sourceMap[sourceNames[i]] = sourceData[i];
  }

  // Update **only columns with matching headers**
  for (let tCol = 0; tCol < targetHeaders.length; tCol++) {
    const header = targetHeaders[tCol];
    const sCol = sourceHeaders.indexOf(header);
    if (sCol === -1) continue; // skip columns not found in source (preserves formulas)

    const colValues = [];
    for (let r = 0; r < targetNames.length; r++) {
      const name = targetNames[r];
      if (sourceMap[name] && sourceMap[name][sCol] !== "") {
        colValues.push([sourceMap[name][sCol]]);
      } else {
        // keep existing value
        colValues.push([targetSheet.getRange(r + 2, tCol + 1).getValue()]);
      }
    }

    // Write only this column in bulk
    targetSheet.getRange(2, tCol + 1, colValues.length, 1).setValues(colValues);
  }

  return "Scores updated from " + sheetName + " (formulas preserved).";
}


// Return all sheet names for dropdown
function getAllSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}
