/**
 * @OnlyCurrentDoc
 */

const MAX_SELECTION_ROWS = 50;

function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Google Sheets Sidebar');
  SpreadsheetApp.getUi().showSidebar(html);
}

function successResult(message, details, warnings) {
  return {
    ok: true,
    message,
    details: details || {},
    warnings: warnings || [],
  };
}

function errorResult(message) {
  return { ok: false, message, details: {}, warnings: [] };
}

/** Select exactly two columns, extending each selection by up to 50 rows. */
function selectColumnsDown() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRangeList();

  if (!selection || selection.getRanges().length !== 2) {
    return errorResult('Select exactly two cells or ranges before continuing.');
  }

  const newRanges = selection.getRanges().map(range => {
    let startRow = range.getRow();
    if (startRow === 1) startRow = 2;

    const rowCount = Math.min(
      MAX_SELECTION_ROWS,
      sheet.getMaxRows() - startRow + 1,
    );
    return sheet.getRange(startRow, range.getColumn(), rowCount, 1);
  });

  sheet.getRangeList(newRanges.map(range => range.getA1Notation())).activate();
  SpreadsheetApp.flush();
  return successResult('Selection is ready for Grade Grabber.');
}

/** Fill the active column with formulas averaging the two highest prior scores. */
function fillAverageFormulas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const targetCol = sheet.getActiveCell().getColumn();

  if (targetCol <= 4) {
    return errorResult(
      'Select column 5 or later so four preceding score columns are available.',
    );
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return errorResult('No roster rows were found below the header row.');
  }

  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const startCol = targetCol - 4;
  const template =
    '=IF(COUNTA(%range%)=0,"",IFERROR(AVERAGE(LARGE(%range%,{1}),LARGE(%range%,{2})),MAX(%range%)))';

  const formulas = names.map((row, index) => {
    if (!row[0]) return [''];
    const rangeA1 = sheet.getRange(index + 2, startCol, 1, 4).getA1Notation();
    return [template.replace(/%range%/g, rangeA1)];
  });

  sheet.getRange(2, targetCol, formulas.length, 1).setFormulas(formulas);
  return successResult('Average formulas were filled successfully.', {
    rowsProcessed: formulas.length,
  });
}

function normalizeMatchValue(value) {
  return String(value == null ? '' : value).trim().toLocaleLowerCase();
}

function groupRowIndexesByName(rows) {
  const groups = new Map();
  rows.forEach((row, index) => {
    const key = normalizeMatchValue(row[0]);
    if (!key) return;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(index);
  });
  return groups;
}

/**
 * Match rows by the case-insensitive value in column A. Duplicate names are
 * paired by occurrence order: first-to-first, second-to-second, and so on.
 */
function buildRowMatches(sourceRows, targetRows) {
  const sourceGroups = groupRowIndexesByName(sourceRows);
  const targetGroups = groupRowIndexesByName(targetRows);
  const matches = new Array(targetRows.length).fill(-1);
  const duplicateNames = [];
  const countMismatches = [];
  const namedTargetCount = Array.from(targetGroups.values()).reduce(
    (total, indexes) => total + indexes.length,
    0,
  );
  let unmatchedTargets = targetRows.length - namedTargetCount;

  targetGroups.forEach((targetIndexes, key) => {
    const sourceIndexes = sourceGroups.get(key) || [];
    const pairCount = Math.min(sourceIndexes.length, targetIndexes.length);
    for (let index = 0; index < pairCount; index++) {
      matches[targetIndexes[index]] = sourceIndexes[index];
    }
    unmatchedTargets += targetIndexes.length - pairCount;

    if (targetIndexes.length > 1 || sourceIndexes.length > 1) {
      duplicateNames.push(String(targetRows[targetIndexes[0]][0]).trim());
    }
    if (sourceIndexes.length !== targetIndexes.length) {
      countMismatches.push({
        name: String(targetRows[targetIndexes[0]][0]).trim(),
        sourceCount: sourceIndexes.length,
        targetCount: targetIndexes.length,
      });
    }
  });

  return { matches, duplicateNames, countMismatches, unmatchedTargets };
}

function buildHeaderLookup(headers) {
  const lookup = new Map();
  const duplicates = [];
  headers.forEach((header, index) => {
    const key = normalizeMatchValue(header);
    if (!key) return;
    if (lookup.has(key)) {
      duplicates.push(String(header).trim());
      return;
    }
    lookup.set(key, index);
  });
  return { lookup, duplicates };
}

/** Update matching columns and rows from another tab in the current file. */
function updateScoresFromSourceSheet(sheetName) {
  if (!sheetName) return errorResult('Select a source sheet to update from.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sheetName);
  const targetSheet = ss.getActiveSheet();

  if (!sourceSheet) return errorResult('The selected source sheet was not found.');
  if (sourceSheet.getSheetId() === targetSheet.getSheetId()) {
    return errorResult('Choose a source sheet other than the active sheet.');
  }

  const sourceLastCol = sourceSheet.getLastColumn();
  const sourceLastRow = sourceSheet.getLastRow();
  const targetLastCol = targetSheet.getLastColumn();
  const targetLastRow = targetSheet.getLastRow();

  if (sourceLastCol < 1 || sourceLastRow < 2) {
    return errorResult('The source sheet is empty or missing roster rows.');
  }
  if (targetLastCol < 1 || targetLastRow < 2) {
    return errorResult('The active sheet has no roster rows to update.');
  }

  const sourceHeaders = sourceSheet
    .getRange(1, 1, 1, sourceLastCol)
    .getValues()[0];
  const sourceData = sourceSheet
    .getRange(2, 1, sourceLastRow - 1, sourceLastCol)
    .getValues();
  const targetHeaders = targetSheet
    .getRange(1, 1, 1, targetLastCol)
    .getValues()[0];
  const targetRange = targetSheet.getRange(2, 1, targetLastRow - 1, targetLastCol);
  const targetData = targetRange.getValues();
  const targetFormulas = targetRange.getFormulas();

  const rowMatch = buildRowMatches(sourceData, targetData);
  const sourceHeader = buildHeaderLookup(sourceHeaders);
  const targetHeader = buildHeaderLookup(targetHeaders);
  const matchedColumns = [];

  targetHeaders.forEach((header, targetColumn) => {
    const headerKey = normalizeMatchValue(header);
    if (!headerKey || !sourceHeader.lookup.has(headerKey)) return;

    const sourceColumn = sourceHeader.lookup.get(headerKey);
    const columnValues = targetData.map((targetRow, targetRowIndex) => {
      const sourceRowIndex = rowMatch.matches[targetRowIndex];
      const existingValue =
        targetFormulas[targetRowIndex][targetColumn] || targetRow[targetColumn];
      if (sourceRowIndex === -1) return [existingValue];

      const sourceValue = sourceData[sourceRowIndex][sourceColumn];
      return [sourceValue === '' ? existingValue : sourceValue];
    });

    targetSheet
      .getRange(2, targetColumn + 1, columnValues.length, 1)
      .setValues(columnValues);
    matchedColumns.push(String(header).trim());
  });

  if (!matchedColumns.length) {
    return errorResult('No matching column headers were found.');
  }

  const warnings = [];
  if (rowMatch.duplicateNames.length) {
    warnings.push(
      'Repeated names were matched by roster order: ' +
        rowMatch.duplicateNames.join(', ') +
        '.',
    );
  }
  rowMatch.countMismatches.forEach(item => {
    warnings.push(
      `${item.name} appears ${item.sourceCount} time(s) in the source and ` +
        `${item.targetCount} time(s) in the active sheet. Unpaired rows were left unchanged.`,
    );
  });
  if (sourceHeader.duplicates.length) {
    warnings.push(
      'Repeated source headers used their first occurrence: ' +
        sourceHeader.duplicates.join(', ') +
        '.',
    );
  }
  if (targetHeader.duplicates.length) {
    warnings.push(
      'Repeated active-sheet headers were all updated from the first matching source header: ' +
        targetHeader.duplicates.join(', ') +
        '.',
    );
  }

  return successResult(`Updated from ${sheetName}.`, {
    matchedRows: rowMatch.matches.filter(index => index !== -1).length,
    unmatchedRows: rowMatch.unmatchedTargets,
    matchedColumns,
    duplicateNames: rowMatch.duplicateNames,
  }, warnings);
}

function getAvailableSourceSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheetId = ss.getActiveSheet().getSheetId();
  return ss
    .getSheets()
    .filter(sheet => sheet.getSheetId() !== activeSheetId)
    .map(sheet => sheet.getName());
}
