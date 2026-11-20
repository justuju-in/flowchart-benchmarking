const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function environment() {
  const tabs = {};
  const sheet = rows => ({
    rows,
    getLastRow() { return this.rows.length; },
    getDataRange() { return {getValues: () => this.rows.map(row => row.slice())}; },
    appendRow(row) { this.rows.push(Array.from(row)); },
    getRange(row, col, height = 1) { return {
      getValue: () => (this.rows[row - 1] || [])[col - 1] || '',
      setValue: value => { this.rows[row - 1] ||= []; this.rows[row - 1][col - 1] = value; },
      getValues: () => this.rows.slice(row - 1, row - 1 + height).map(r => [r[col - 1]]),
      setValues: values => values.forEach((v, i) => { this.rows[row - 1 + i] = Array.from(v); })
    }; },
    setFrozenRows() {}
  });
  tabs.BenchmarkSetA = sheet([['header'], ...Array.from({length: 8}, (_, i) => ['', '', '', `FCP00${i + 1}`, `image${i}`, '', ''])]);
  tabs.BenchmarkRuns = sheet([['Run ID', 'Set Name', 'Model', 'Reasoning Level', 'Status', 'Started At', 'Finished At']]);
  tabs.BenchmarkResults = sheet([['Run ID']]);
  const ss = {getSheetByName: name => tabs[name] || null, insertSheet: name => (tabs[name] = sheet([]))};
  const context = vm.createContext({SpreadsheetApp: {openById: () => ss}});
  vm.runInContext(fs.readFileSync('benchmarkRuns.js', 'utf8'), context);
  return {context, tabs};
}

test('new pilot queues exactly the chosen five items', () => {
  const {context, tabs} = environment();
  context.createConfiguredBenchmarkRun('BenchmarkSetA', 'model', 'medium', 5);
  context.queueBenchmarkRun();
  assert.equal(tabs.BenchmarkResults.rows.length - 1, 5);
  assert.equal(tabs.BenchmarkRuns.rows[1][7], 5);
  assert.equal(tabs.BenchmarkResults.rows[5][2], 6);
});

test('sample size rejects out-of-range values', () => {
  const {context} = environment();
  assert.throws(() => context.createConfiguredBenchmarkRun('BenchmarkSetA', 'model', 'medium', 11), /between 1 and 10/);
});

test('comparison preserves existing report and refuses different sets', () => {
  const {context, tabs} = environment();
  const result = (run, set, verdict) => [run, set, 2, 'FCP001', '', '', '', '', '', '', '', '', verdict];
  tabs.BenchmarkResults.rows.push(result('RUN-001', 'BenchmarkSetA', 'AC'), result('RUN-002', 'BenchmarkSetA', 'WA'));
  context.buildBenchmarkRunComparisonReport('RUN-001', 'RUN-002');
  const report = tabs['BenchmarkComparison_RUN-001_RUN-002'];
  assert.equal(report.rows[1][6], 'AC→WA');
  assert.throws(() => context.buildBenchmarkRunComparisonReport('RUN-001', 'RUN-002'), /already exists/);
  tabs.BenchmarkResults.rows.push(result('RUN-003', 'BenchmarkSetB', 'AC'));
  assert.throws(() => context.buildBenchmarkRunComparisonReport('RUN-001', 'RUN-003'), /same benchmark set/);
});
