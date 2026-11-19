const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const context = vm.createContext({});
vm.runInContext(fs.readFileSync('benchmarkSets.js', 'utf8'), context);

test('Set B combines repeated flowcharts and normalizes problem codes', () => {
  const headers = ['flowchart', 'f_id', 'issue_reported', 'rca', 'debugging_status', 'debugging_comments'];
  const link = 'https://drive.google.com/open?id=1234567890123456789012345';
  const items = context.benchmarkBuildSetBItems_([headers,
    [link, 'FCP23', 'wrong output', 'condition', '', ''],
    [link, 'FCP23', 'wrong output', 'condition', 'reviewed', ''],
    ['', 'FCP24', '', '', '', '']
  ]);
  assert.equal(items.length, 1);
  assert.equal(items[0].problemCode, 'FCP023');
  assert.equal(items[0].issueReported, 'wrong output');
  assert.equal(items[0].debuggingStatus, 'reviewed');
});

test('Set B resolves numeric f_id rows through the historical flowchart URL', () => {
  const flaggedHeaders = ['flowchart', 'f_id', 'issue_reported', 'rca', 'debugging_status', 'debugging_comments'];
  const sourceHeaders = ['submitted_at', 'problem_id', 'image_url'];
  const link = 'https://drive.google.com/open?id=1234567890123456789012345';
  const mapping = context.benchmarkProblemCodesByFile_([
    sourceHeaders,
    ['2025-07-20', 'FCP052 Print Elements from Hardcoded Array', link]
  ]);
  const items = context.benchmarkBuildSetBItems_([
    flaggedHeaders,
    [link, '351', 'wrong output', '', '', '']
  ], mapping);
  assert.equal(items.length, 1);
  assert.equal(items[0].sourceReference, '351');
  assert.equal(items[0].problemCode, 'FCP052');
});

test('Set B rejects missing source columns', () => {
  assert.throws(() => context.benchmarkBuildSetBItems_([['flowchart'], ['x']]), /missing column/);
});
