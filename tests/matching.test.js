'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync('shortcuts_sidebar.gs', 'utf8');
const context = { console };
vm.createContext(context);
vm.runInContext(
  `${source}\nthis.testApi = { normalizeMatchValue, buildRowMatches, buildHeaderLookup, buildRosterDiff, shiftFormulaA1 };`,
  context,
);
const api = context.testApi;

test('normalizes names and headers case-insensitively', () => {
  assert.equal(api.normalizeMatchValue('  Kevin CASSIDY  '), 'kevin cassidy');
  assert.equal(api.normalizeMatchValue(null), '');
});

test('matches repeated names by occurrence order', () => {
  const result = api.buildRowMatches(
    [
      ['Alex Smith', 91],
      ['Jordan Lee', 84],
      [' alex smith ', 96],
    ],
    [
      ['ALEX SMITH'],
      ['Alex Smith'],
      ['Jordan Lee'],
    ],
  );

  assert.deepEqual(Array.from(result.matches), [0, 2, 1]);
  assert.deepEqual(Array.from(result.duplicateNames), ['ALEX SMITH']);
  assert.equal(result.unmatchedTargets, 0);
});

test('leaves extra duplicate targets unmatched and reports count mismatch', () => {
  const result = api.buildRowMatches(
    [['Alex Smith']],
    [['Alex Smith'], ['alex smith']],
  );

  assert.deepEqual(Array.from(result.matches), [0, -1]);
  assert.equal(result.unmatchedTargets, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(result.countMismatches)), [
    { name: 'Alex Smith', sourceCount: 1, targetCount: 2 },
  ]);
});

test('counts blank destination names as unmatched rows', () => {
  const result = api.buildRowMatches(
    [['Alex Smith']],
    [['Alex Smith'], ['', 'orphaned score']],
  );

  assert.deepEqual(Array.from(result.matches), [0, -1]);
  assert.equal(result.unmatchedTargets, 1);
});

test('uses the first duplicate source header and reports the duplicate', () => {
  const result = api.buildHeaderLookup(['Name', 'Quiz 1', ' quiz 1 ']);

  assert.equal(result.lookup.get('quiz 1'), 1);
  assert.deepEqual(Array.from(result.duplicates), ['quiz 1']);
});

test('identifies incoming and departing roster occurrences', () => {
  const result = api.buildRosterDiff(
    [['Alex'], ['Sam'], ['Sam'], ['New Student']],
    [['Alex'], ['Sam'], ['Former Student']],
  );
  assert.deepEqual(JSON.parse(JSON.stringify(result)), {
    incoming: [{ sourceIndex: 2, name: 'Sam' }, { sourceIndex: 3, name: 'New Student' }],
    departures: [{ row: 4, name: 'Former Student' }],
  });
});

test('shifts relative formula references from the example destination', () => {
  assert.equal(api.shiftFormulaA1('=MAX(B2:C2)', 'D2', 2, 7), '=MAX(E2:F2)');
  assert.equal(api.shiftFormulaA1('=MAX(B2:C2)', 'D2', 5, 7), '=MAX(E5:F5)');
});

test('preserves absolute references, sheet names, and quoted strings', () => {
  assert.equal(
    api.shiftFormulaA1('=IF(B2="A2",\'Scores 1\'!$C2+$A$1)', 'D2', 4, 6),
    '=IF(D4="A2",\'Scores 1\'!$C4+$A$1)',
  );
});
