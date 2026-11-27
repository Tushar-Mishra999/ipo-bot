const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const samples = JSON.parse(fs.readFileSync('tests-samples.json', 'utf8'));
const elements = new Proxy({}, {get(target, key) {
  return target[key] ||= {addEventListener() {}, querySelectorAll() {return [];}, set innerHTML(value) {}, get innerHTML() {return '';}};
}});
const context = vm.createContext({
  console, setTimeout, clearTimeout,
  window: {IPO_AI_ENDPOINT: ''},
  document: {getElementById: id => elements[id], querySelectorAll: () => []},
  fetch: async () => ({ok: true, json: async () => ({preferences: {default: 5, min: 0, max: 10, step: 1,
    holding: {label: '', info: '', low: '', high: ''}, safety: {label: '', info: '', low: '', high: ''},
    income: {label: '', info: '', low: '', high: ''}}})}),
  XLSX: {utils: {sheet_to_json: sheet => sheet.rows}},
});
vm.runInContext(fs.readFileSync('ipo_bot/analysis.js', 'utf8'), context);
const analyze = rows => vm.runInContext('calculate(parseWorkbook(workbook))', Object.assign(context, {
  workbook: {Sheets: {'Data Sheet': {rows}}}
}));

for (const [file, reference] of [['Crystal Crop.xlsx', -2.68], ['Mahanadi Coalfld.xlsx', -2.36]]) {
  test(`${file}: six annual years and reference M-Score`, () => {
    const result = analyze(samples[file]);
    assert.ok(Number.isFinite(result.metrics['Beneish M-Score']));
    assert.equal(result.diagnostics['Beneish M-Score'].skipReason, undefined);
    assert.ok(Math.abs(result.metrics['Beneish M-Score'] - reference) <= 0.05,
      `${result.metrics['Beneish M-Score']} differs from ${reference}`);
  });
}

test('six years with blank optional rows remain calculable', () => {
  const rows = structuredClone(samples['Crystal Crop.xlsx']);
  for (const row of rows) if (['Power and Fuel', 'Dividend Amount', 'Capital Work in Progress'].includes(row[0])) row.fill(null, 1);
  const result = analyze(rows);
  assert.ok(Number.isFinite(result.metrics['Beneish M-Score']));
});

test('two years with early blank columns use the last two Report Date years', () => {
  const rows = structuredClone(samples['Crystal Crop.xlsx']);
  for (const row of rows) if (row[0] === 'Report Date') for (let i = 5; i < 9; i++) row[i] = null;
  const result = analyze(rows);
  assert.ok(Number.isFinite(result.metrics['Beneish M-Score']));
  assert.match(result.diagnostics['Sales CAGR'].notes.join(' '), /1-year growth/);
});

test('missing core row gives an exact skip reason', () => {
  const rows = structuredClone(samples['Crystal Crop.xlsx']).filter(row => row[0] !== 'Receivables');
  const result = analyze(rows);
  assert.equal(result.metrics['Beneish M-Score'], null);
  assert.match(result.diagnostics['Beneish M-Score'].skipReason, /Receivables row missing/);
});

test('zero prior receivables and SG&A are neutral and explained', () => {
  const rows = structuredClone(samples['Crystal Crop.xlsx']);
  for (const row of rows) if (['Receivables', 'Selling and admin'].includes(row[0])) row[9] = 0;
  const result = analyze(rows);
  assert.ok(Number.isFinite(result.metrics['Beneish M-Score']));
  assert.match(result.diagnostics['Beneish M-Score'].notes.join(' '), /DSRI set to 1/);
  assert.match(result.diagnostics['Beneish M-Score'].notes.join(' '), /SGAI set to 1/);
});

test('an extreme index is capped with a named warning', () => {
  const rows = structuredClone(samples['Crystal Crop.xlsx']);
  for (const row of rows) if (row[0] === 'Receivables') row[9] = 0.01;
  const result = analyze(rows);
  assert.match(result.diagnostics['Beneish M-Score'].notes.join(' '), /DSRI capped at 3/);
});
