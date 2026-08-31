'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync('shortcuts_sidebar.gs', 'utf8');
const context = { console };
vm.createContext(context);
vm.runInContext(
  `${source}\nthis.testApi = { normalizeMatchValue, buildRowMatches, buildHeaderLookup };`,
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
