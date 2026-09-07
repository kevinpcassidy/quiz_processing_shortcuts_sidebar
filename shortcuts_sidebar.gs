/**
 * @OnlyCurrentDoc
 */

const CUSTOM_FORMULAS_PROPERTY = 'spreadsheetSidekick.customFormulas.v1';
const BUILT_IN_FORMULAS = [{
  id: 'two-highest-average',
  name: 'Average of the Two Highest out of Previous Four Cells',
  formula: '=IF(COUNTA(A2:D2)=0,"",IFERROR(AVERAGE(LARGE(A2:D2,{1}),LARGE(A2:D2,{2})),MAX(A2:D2)))',
  exampleDestination: 'E2',
  builtIn: true,
}];

function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Open Spreadsheet Sidekick', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Spreadsheet Sidekick');
  SpreadsheetApp.getUi().showSidebar(html);
}

function successResult(message, details, warnings) {
  return { ok: true, message, details: details || {}, warnings: warnings || [] };
}

function errorResult(message) {
  return { ok: false, message, details: {}, warnings: [] };
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

/** Duplicate names are paired first-to-first, second-to-second, and so on. */
function buildRowMatches(sourceRows, targetRows) {
  const sourceGroups = groupRowIndexesByName(sourceRows);
  const targetGroups = groupRowIndexesByName(targetRows);
  const matches = new Array(targetRows.length).fill(-1);
  const duplicateNames = [];
  const countMismatches = [];
  let unmatchedTargets = targetRows.filter(row => !normalizeMatchValue(row[0])).length;

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

function occurrenceTokens(rows) {
  const counts = new Map();
  return rows.map(row => {
    const key = normalizeMatchValue(row[0]);
    if (!key) return '';
    const occurrence = (counts.get(key) || 0) + 1;
    counts.set(key, occurrence);
    return `${key}\u0000${occurrence}`;
  });
}

function buildRosterDiff(sourceRows, targetRows) {
  const sourceTokens = occurrenceTokens(sourceRows);
  const targetTokens = occurrenceTokens(targetRows);
  const sourceSet = new Set(sourceTokens.filter(Boolean));
  const targetSet = new Set(targetTokens.filter(Boolean));
  return {
    incoming: sourceRows.reduce((items, row, index) => {
      if (sourceTokens[index] && !targetSet.has(sourceTokens[index])) {
        items.push({ sourceIndex: index, name: String(row[0]).trim() });
      }
      return items;
    }, []),
    departures: targetRows.reduce((items, row, index) => {
      if (targetTokens[index] && !sourceSet.has(targetTokens[index])) {
        items.push({ row: index + 2, name: String(row[0]).trim() });
      }
      return items;
    }, []),
  };
}

function getSheetPair(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sheetName);
  const targetSheet = ss.getActiveSheet();
  if (!sourceSheet) return { error: errorResult('The selected source sheet was not found.') };
  if (sourceSheet.getSheetId() === targetSheet.getSheetId()) {
    return { error: errorResult('Choose a source sheet other than the active sheet.') };
  }
  if (sourceSheet.getLastColumn() < 1 || sourceSheet.getLastRow() < 2) {
    return { error: errorResult('The source sheet is empty or missing roster rows.') };
  }
  if (targetSheet.getLastColumn() < 1) {
    return { error: errorResult('The destination sheet needs a header row before it can be updated.') };
  }
  return { sourceSheet, targetSheet };
}

function readRows(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
}

function previewSourceSheetUpdate(sheetName) {
  if (!sheetName) return errorResult('Choose a source sheet first.');
  const pair = getSheetPair(sheetName);
  if (pair.error) return pair.error;
  const diff = buildRosterDiff(readRows(pair.sourceSheet), readRows(pair.targetSheet));
  return successResult('Roster comparison is ready.', {
    sourceSheet: pair.sourceSheet.getName(),
    destinationSheet: pair.targetSheet.getName(),
    incoming: diff.incoming,
    departures: diff.departures,
    requiresIncomingConfirmation: diff.incoming.length > 5,
  });
}

/** Plan insertions in their eventual destination order without rereading a sheet. */
function buildInsertionGroups(sourceRows, targetRows) {
  const sourceTokens = occurrenceTokens(sourceRows);
  const workingTargetTokens = occurrenceTokens(targetRows);
  const insertions = [];

  sourceTokens.forEach((token, sourceIndex) => {
    if (!token || workingTargetTokens.includes(token)) return;
    let targetIndex = workingTargetTokens.length;
    for (let next = sourceIndex + 1; next < sourceTokens.length; next++) {
      const nextTargetIndex = workingTargetTokens.indexOf(sourceTokens[next]);
      if (nextTargetIndex !== -1) {
        targetIndex = nextTargetIndex;
        break;
      }
    }
    insertions.push({ sourceIndex, targetIndex });
    workingTargetTokens.splice(targetIndex, 0, token);
  });

  return insertions.reduce((groups, insertion) => {
    const previous = groups[groups.length - 1];
    if (previous &&
        insertion.sourceIndex === previous.sourceIndexes[previous.sourceIndexes.length - 1] + 1 &&
        insertion.targetIndex === previous.targetIndex + previous.sourceIndexes.length) {
      previous.sourceIndexes.push(insertion.sourceIndex);
    } else {
      groups.push({
        targetIndex: insertion.targetIndex,
        sourceIndexes: [insertion.sourceIndex],
      });
    }
    return groups;
  }, []);
}

function getNeighborFormulaTemplates(sheet, insertionRow, rowCount, lastColumn) {
  const formulasAbove = insertionRow > 2
    ? sheet.getRange(insertionRow - 1, 1, 1, lastColumn).getFormulas()[0]
    : new Array(lastColumn).fill('');
  const rowBelow = insertionRow + rowCount;
  const formulasBelow = rowBelow <= sheet.getLastRow()
    ? sheet.getRange(rowBelow, 1, 1, lastColumn).getFormulas()[0]
    : new Array(lastColumn).fill('');
  return formulasAbove.map((formula, index) => ({
    formula: formula || formulasBelow[index],
    templateRow: formula ? insertionRow - 1 : rowBelow,
  }));
}

function insertMissingSourceRows(sourceSheet, targetSheet) {
  const sourceLastColumn = sourceSheet.getLastColumn();
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceLastColumn).getValues()[0];
  const sourceRows = readRows(sourceSheet);
  const targetRows = readRows(targetSheet);
  const groups = buildInsertionGroups(sourceRows, targetRows);
  const lastColumn = targetSheet.getLastColumn();
  const targetHeaders = targetSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const sourceLookup = buildHeaderLookup(sourceHeaders).lookup;

  groups.forEach(group => {
    const insertionRow = group.targetIndex + 2;
    const rowCount = group.sourceIndexes.length;
    const lastDataRow = targetSheet.getLastRow();
    if (insertionRow <= lastDataRow) {
      targetSheet.insertRowsBefore(insertionRow, rowCount);
    } else {
      const requiredLastRow = insertionRow + rowCount - 1;
      if (requiredLastRow > targetSheet.getMaxRows()) {
        targetSheet.insertRowsAfter(
          targetSheet.getMaxRows(),
          requiredLastRow - targetSheet.getMaxRows(),
        );
      }
    }

    const formulaTemplates = getNeighborFormulaTemplates(
      targetSheet,
      insertionRow,
      rowCount,
      lastColumn,
    );
    const values = group.sourceIndexes.map((sourceIndex, offset) =>
      targetHeaders.map((header, targetColumn) => {
        const sourceColumn = sourceLookup.get(normalizeMatchValue(header));
        if (sourceColumn !== undefined) return sourceRows[sourceIndex][sourceColumn];
        const template = formulaTemplates[targetColumn];
        if (!template.formula) return '';
        return shiftFormulaA1(
          template.formula,
          `${columnLetters(targetColumn + 1)}${template.templateRow}`,
          insertionRow + offset,
          targetColumn + 1,
        );
      }),
    );
    targetSheet.getRange(insertionRow, 1, rowCount, lastColumn).setValues(values);
  });
}

function updateScoresFromSourceSheet(sheetName, departureRows) {
  const pair = getSheetPair(sheetName);
  if (pair.error) return pair.error;
  const sourceSheet = pair.sourceSheet;
  const targetSheet = pair.targetSheet;
  const originalDiff = buildRosterDiff(readRows(sourceSheet), readRows(targetSheet));
  const allowedDepartures = new Map(originalDiff.departures.map(item => [item.row, item.name]));
  const rowsToDelete = (departureRows || []).filter(item =>
    allowedDepartures.get(Number(item.row)) === item.name,
  );
  rowsToDelete.sort((a, b) => b.row - a.row).forEach(item => targetSheet.deleteRow(Number(item.row)));

  insertMissingSourceRows(sourceSheet, targetSheet);

  const sourceLastCol = sourceSheet.getLastColumn();
  const sourceLastRow = sourceSheet.getLastRow();
  const targetLastCol = targetSheet.getLastColumn();
  const targetLastRow = targetSheet.getLastRow();
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceLastCol).getValues()[0];
  const sourceData = sourceSheet.getRange(2, 1, sourceLastRow - 1, sourceLastCol).getValues();
  const targetHeaders = targetSheet.getRange(1, 1, 1, targetLastCol).getValues()[0];
  const targetRange = targetSheet.getRange(2, 1, targetLastRow - 1, targetLastCol);
  const targetData = targetRange.getValues();
  const targetFormulas = targetRange.getFormulas();
  const rowMatch = buildRowMatches(sourceData, targetData);
  const sourceHeader = buildHeaderLookup(sourceHeaders);
  const targetHeader = buildHeaderLookup(targetHeaders);
  const matchedColumns = [];

  targetHeaders.forEach((header, targetColumn) => {
    const sourceColumn = sourceHeader.lookup.get(normalizeMatchValue(header));
    if (sourceColumn === undefined) return;
    const columnValues = targetData.map((targetRow, rowIndex) => {
      const sourceRowIndex = rowMatch.matches[rowIndex];
      const existing = targetFormulas[rowIndex][targetColumn] || targetRow[targetColumn];
      if (sourceRowIndex === -1) return [existing];
      const sourceValue = sourceData[sourceRowIndex][sourceColumn];
      return [sourceValue === '' ? existing : sourceValue];
    });
    targetSheet.getRange(2, targetColumn + 1, columnValues.length, 1).setValues(columnValues);
    matchedColumns.push(String(header).trim());
  });

  if (!matchedColumns.length) return errorResult('No matching column headers were found.');
  const warnings = [];
  const keptDepartures = originalDiff.departures.length - rowsToDelete.length;
  if (keptDepartures) warnings.push(`${keptDepartures} destination-only student row(s) were kept.`);
  if (rowMatch.duplicateNames.length) {
    warnings.push(`Repeated names were matched by roster order: ${rowMatch.duplicateNames.join(', ')}.`);
  }
  if (sourceHeader.duplicates.length) {
    warnings.push(`Repeated source headers used their first occurrence: ${sourceHeader.duplicates.join(', ')}.`);
  }
  if (targetHeader.duplicates.length) {
    warnings.push(`Repeated destination headers were updated from the first matching source header: ${targetHeader.duplicates.join(', ')}.`);
  }
  return successResult(`Updated from ${sheetName}.`, {
    matchedRows: rowMatch.matches.filter(index => index !== -1).length,
    unmatchedRows: rowMatch.unmatchedTargets,
    addedRows: originalDiff.incoming.length,
    deletedRows: rowsToDelete.length,
    matchedColumns,
  }, warnings);
}

function buildSourceSheetInfo(sheetNames, activeSheetName) {
  return {
    activeSheetName,
    names: sheetNames.filter(name => name !== activeSheetName),
  };
}

function getAvailableSourceSheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return buildSourceSheetInfo(
    ss.getSheets().map(sheet => sheet.getName()),
    ss.getActiveSheet().getName(),
  );
}

function parseCellAddress(address) {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(String(address || '').trim());
  if (!match) return null;
  let column = 0;
  for (const letter of match[1].toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64;
  return { column, row: Number(match[2]) };
}

function columnLetters(column) {
  let result = '';
  while (column > 0) {
    column--;
    result = String.fromCharCode(65 + column % 26) + result;
    column = Math.floor(column / 26);
  }
  return result;
}

/** Shift A1 references while leaving references inside quoted strings untouched. */
function shiftFormulaA1(formula, anchorAddress, targetRow, targetColumn) {
  const anchor = parseCellAddress(anchorAddress);
  if (!anchor) throw new Error('The example destination must be a cell such as E2.');
  const rowDelta = targetRow - anchor.row;
  const columnDelta = targetColumn - anchor.column;
  const parts = String(formula).split(/("(?:[^"]|"")*")/g);
  return parts.map((part, index) => {
    if (index % 2) return part;
    return part.replace(/(^|[^A-Z0-9_])((?:'(?:[^']|'')+'|[A-Z0-9_ ]+)!|)(\$?)([A-Z]{1,3})(\$?)(\d+)/gi,
      (whole, prefix, sheet, absoluteColumn, letters, absoluteRow, rowText) => {
        const originalColumn = parseCellAddress(`${letters}${rowText}`).column;
        const shiftedColumn = absoluteColumn ? originalColumn : originalColumn + columnDelta;
        const shiftedRow = absoluteRow ? Number(rowText) : Number(rowText) + rowDelta;
        if (shiftedColumn < 1 || shiftedRow < 1) throw new Error('The formula shifts outside the sheet.');
        return `${prefix}${sheet}${absoluteColumn}${columnLetters(shiftedColumn)}${absoluteRow}${shiftedRow}`;
      });
  }).join('');
}

function readCustomFormulas() {
  const raw = PropertiesService.getUserProperties().getProperty(CUSTOM_FORMULAS_PROPERTY);
  if (!raw) return [];
  try {
    const items = JSON.parse(raw);
    return Array.isArray(items) ? items : [];
  } catch (error) {
    return [];
  }
}

function getFormulaLibrary() {
  return BUILT_IN_FORMULAS.concat(readCustomFormulas());
}

function saveCustomFormula(item) {
  const name = String(item && item.name || '').trim();
  const formula = String(item && item.formula || '').trim();
  const exampleDestination = String(item && item.exampleDestination || '').trim().toUpperCase();
  if (!name) return errorResult('Enter a formula name.');
  if (!formula.startsWith('=')) return errorResult('The formula must begin with =.');
  const anchor = parseCellAddress(exampleDestination);
  if (!anchor || anchor.row !== 2) return errorResult('Use a row 2 cell, such as E2, for the example destination.');
  const formulas = readCustomFormulas();
  const id = String(item.id || `custom-${Date.now()}`);
  const duplicate = getFormulaLibrary().find(entry =>
    entry.name.toLocaleLowerCase() === name.toLocaleLowerCase() && entry.id !== id,
  );
  if (duplicate) return errorResult('A formula with that name already exists.');
  const saved = { id, name, formula, exampleDestination, builtIn: false };
  const index = formulas.findIndex(entry => entry.id === id);
  if (index === -1) formulas.push(saved); else formulas[index] = saved;
  PropertiesService.getUserProperties().setProperty(CUSTOM_FORMULAS_PROPERTY, JSON.stringify(formulas));
  return successResult('Formula saved.', { formulas: getFormulaLibrary(), selectedId: id });
}

function deleteCustomFormula(id) {
  const formulas = readCustomFormulas();
  const next = formulas.filter(item => item.id !== id);
  if (next.length === formulas.length) return errorResult('That custom formula was not found.');
  PropertiesService.getUserProperties().setProperty(CUSTOM_FORMULAS_PROPERTY, JSON.stringify(next));
  return successResult('Formula deleted.', { formulas: getFormulaLibrary() });
}

function findFormula(id) {
  return getFormulaLibrary().find(item => item.id === id);
}

function previewFormulaFill(id) {
  const item = findFormula(id);
  if (!item) return errorResult('Choose a formula first.');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const targetColumn = sheet.getActiveCell().getColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return errorResult('No roster rows were found below the header.');
  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let finalNameIndex = -1;
  names.forEach((row, index) => { if (normalizeMatchValue(row[0])) finalNameIndex = index; });
  if (finalNameIndex === -1) return errorResult('No names were found in column A.');
  const values = sheet.getRange(2, targetColumn, finalNameIndex + 1, 1).getDisplayValues();
  const formulas = sheet.getRange(2, targetColumn, finalNameIndex + 1, 1).getFormulas();
  const occupiedCells = values.filter((row, index) => row[0] !== '' || formulas[index][0] !== '').length;
  return successResult('Formula fill is ready.', {
    formulaName: item.name,
    column: columnLetters(targetColumn),
    rowCount: finalNameIndex + 1,
    occupiedCells,
  });
}

function fillSavedFormula(id, replaceExisting) {
  const item = findFormula(id);
  if (!item) return errorResult('Choose a formula first.');
  const preview = previewFormulaFill(id);
  if (!preview.ok) return preview;
  if (preview.details.occupiedCells && !replaceExisting) {
    return errorResult('Confirm that existing destination content may be replaced.');
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const targetColumn = sheet.getActiveCell().getColumn();
  const names = sheet.getRange(2, 1, preview.details.rowCount, 1).getValues();
  const formulas = names.map((row, index) => [
    normalizeMatchValue(row[0])
      ? shiftFormulaA1(item.formula, item.exampleDestination, index + 2, targetColumn)
      : '',
  ]);
  sheet.getRange(2, targetColumn, formulas.length, 1).setFormulas(formulas);
  return successResult(`Filled “${item.name}” in column ${columnLetters(targetColumn)}.`, {
    rowsProcessed: formulas.filter(row => row[0]).length,
  });
}
