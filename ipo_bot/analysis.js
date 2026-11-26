const $ = id => document.getElementById(id);
const number = value => {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  const cleaned = String(value ?? '').replace(/,/g, '').replace(/[₹%]/g, '').trim();
  if (!cleaned || cleaned === '-' || cleaned === '--') return null;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
};
const divide = (a, b) => a == null || b == null || b === 0 ? null : a / b;
const add = (...values) => values.every(v => v != null) ? values.reduce((a, b) => a + b, 0) : null;
const difference = (a, b) => a == null || b == null ? null : a - b;
const label = value => String(value ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
const rowNames = {
  sales: ['Sales', 'Revenue'], raw: ['Raw Material Cost'], inventoryChange: ['Change in Inventory'],
  power: ['Power and Fuel'], manufacturing: ['Other Mfr. Exp'], selling: ['Selling and admin'],
  otherIncome: ['Other Income'], depreciation: ['Depreciation'], interest: ['Interest'],
  pbt: ['Profit before tax'], pat: ['Net profit'], dividend: ['Dividend Amount'],
  equity: ['Equity Share Capital'], reserves: ['Reserves'], borrowings: ['Borrowings'],
  liabilities: ['Other Liabilities'], assets: ['Total (assets)', 'Total Assets', 'Total'],
  netBlock: ['Net Block'], cwip: ['Capital Work in Progress'], receivables: ['Receivables'],
  inventory: ['Inventory'], cash: ['Cash & Bank', 'Cash and Bank'],
  cfo: ['Cash from Operating Activity']
};

function parseWorkbook(workbook) {
  const sheet = workbook.Sheets['Data Sheet'];
  if (!sheet) throw new Error('This workbook needs a Data Sheet tab.');
  const rows = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: null, raw: true});
  const indexed = new Map();
  for (const row of rows) {
    const key = label(row[0]);
    if (key && !indexed.has(key)) indexed.set(key, row);
  }
  const dateRows = rows.filter(row => /reportdate|yearended|yearending/i.test(label(row[0])));
  const header = dateRows.find(row => row.slice(1).filter(v => v != null && v !== '').length >= 2) ||
    rows.find(row => row.slice(1).filter(v => v instanceof Date || typeof v === 'number' && v > 30000 && v < 80000).length >= 2);
  if (!header) throw new Error('Could not find the annual year columns in Data Sheet.');
  const columns = header.map((v, i) => i).filter(i => i > 0 && header[i] != null && header[i] !== '' &&
    !/ttm|quarter/i.test(String(header[i])));
  if (!columns.length) throw new Error('No annual data columns were found.');
  const years = columns.map(i => {
    const value = header[i];
    if (value instanceof Date) return String(value.getFullYear());
    if (typeof value === 'number' && value > 30000) return String(new Date(Date.UTC(1899, 11, 30) + value * 86400000).getUTCFullYear());
    const match = String(value).match(/20\d{2}|\d{2}$/);
    return match ? String(match[0]).replace(/^([0-9]{2})$/, '20$1') : String(value);
  });
  const data = {};
  for (const [key, aliases] of Object.entries(rowNames)) {
    const row = aliases.map(name => indexed.get(label(name))).find(Boolean);
    data[key] = columns.map(i => row ? number(row[i]) : null);
  }
  return {years, data};
}

function calculate(parsed) {
  const {years, data} = parsed;
  const n = years.length;
  const at = (key, i = n - 1) => data[key][i];
  const previous = key => at(key, n - 2);
  const latest = key => at(key);
  const currentAssets = i => add(at('receivables', i), at('inventory', i), at('cash', i));
  const ebit = i => add(at('pbt', i), at('interest', i));
  const cogs = i => add(at('raw', i), at('power', i), at('manufacturing', i),
    at('inventoryChange', i) == null ? null : -at('inventoryChange', i));
  const equity = i => add(at('equity', i), at('reserves', i));
  const liabilities = i => add(at('borrowings', i), at('liabilities', i));
  const grossMargin = i => difference(1, divide(cogs(i), at('sales', i)));
  const receivableDays = i => divide(at('receivables', i) == null ? null : at('receivables', i) * 365, at('sales', i));
  const margin = i => divide(add(at('pbt', i), at('interest', i), at('depreciation', i),
    at('otherIncome', i) == null ? null : -at('otherIncome', i)), at('sales', i));
  const metrics = {};
  const validPat = data.pat.map((v, i) => v != null && v > 0 && at('cfo', i) != null ? i : -1).filter(i => i >= 0);
  metrics['CFO / PAT'] = validPat.length ? divide(validPat.reduce((s, i) => s + at('cfo', i), 0), validPat.reduce((s, i) => s + at('pat', i), 0)) : null;
  metrics['Accruals'] = divide(difference(latest('pat'), latest('cfo')), latest('assets'));
  const daysNow = receivableDays(n - 1), daysBefore = receivableDays(0);
  metrics['Receivable days change'] = n > 1 ? divide(difference(daysNow, daysBefore), daysBefore) : null;
  metrics['Other income share'] = divide(latest('otherIncome'), latest('pbt'));
  const z1 = divide(difference(currentAssets(n - 1), latest('liabilities')), latest('assets'));
  const z2 = divide(latest('reserves'), latest('assets'));
  const z3 = divide(ebit(n - 1), latest('assets'));
  const z4 = divide(equity(n - 1), liabilities(n - 1));
  metrics["Altman Z''"] = [z1, z2, z3, z4].every(v => v != null) ? 6.56*z1 + 3.26*z2 + 6.72*z3 + 1.05*z4 : null;
  metrics['Interest cover'] = divide(ebit(n - 1), latest('interest'));
  metrics['Debt / Equity'] = divide(latest('borrowings'), equity(n - 1));
  const cagr = key => n > 1 && at(key, 0) > 0 && latest(key) > 0 ? Math.pow(latest(key)/at(key, 0), 1/(n-1))-1 : null;
  metrics['Sales CAGR'] = cagr('sales');
  metrics['PAT CAGR'] = cagr('pat');
  metrics['ROE'] = divide(latest('pat'), equity(n - 1));
  metrics['EBITDA margin trend'] = n > 1 ? difference(margin(n - 1), margin(0)) : null;
  const capex = n > 1 ? add(difference(latest('netBlock'), previous('netBlock')),
    difference(latest('cwip'), previous('cwip')), latest('depreciation')) : null;
  const fcf = difference(latest('cfo'), capex);
  metrics['FCF / PAT'] = latest('pat') > 0 ? divide(fcf, latest('pat')) : null;
  metrics['Dividend payout vs cash'] = latest('cfo') > 0 && latest('dividend') != null ?
    Math.max(0, 1 - latest('dividend') / latest('cfo')) : null;
  metrics['Cash / Borrowings'] = divide(latest('cash'), latest('borrowings'));
  if (n > 1) {
    const dsri = divide(divide(latest('receivables'), latest('sales')), divide(previous('receivables'), previous('sales')));
    const gmi = divide(grossMargin(n-2), grossMargin(n-1));
    const aqi = divide(difference(1, divide(add(latest('netBlock'), currentAssets(n-1)), latest('assets'))),
      difference(1, divide(add(previous('netBlock'), currentAssets(n-2)), previous('assets'))));
    const sgi = divide(latest('sales'), previous('sales'));
    const depi = divide(divide(previous('depreciation'), add(previous('depreciation'), previous('netBlock'))),
      divide(latest('depreciation'), add(latest('depreciation'), latest('netBlock'))));
    const sgai = divide(divide(latest('selling'), latest('sales')), divide(previous('selling'), previous('sales')));
    const lvgi = divide(divide(liabilities(n-1), latest('assets')), divide(liabilities(n-2), previous('assets')));
    const tata = divide(difference(latest('pat'), latest('cfo')), latest('assets'));
    metrics['Beneish M-Score'] = [dsri,gmi,aqi,sgi,depi,sgai,lvgi,tata].every(v => v != null) ?
      -4.84 + 0.92*dsri + 0.528*gmi + 0.404*aqi + 0.892*sgi + 0.115*depi - 0.172*sgai + 4.679*tata - 0.327*lvgi : null;
  }
  return {metrics, extra: {receivableDays: daysNow, EBITDA: margin(n-1), capex, fcf}, lossMaking: data.pat.some(v => v != null && v <= 0)};
}

function score(result, config) {
  const pillars = [];
  for (const [name, definition] of Object.entries(config.pillars)) {
    const items = Object.entries(definition.metrics).map(([metric, [bad, good]]) => {
      const value = result.metrics[metric];
      const points = value == null ? null : Math.max(0, Math.min(100, 100*(value-bad)/(good-bad)));
      return {metric, value, points};
    });
    const available = items.filter(item => item.points != null);
    pillars.push({name, weight: definition.weight, items,
      coverage: Math.round(100*available.length/items.length),
      score: available.length ? available.reduce((s, item) => s+item.points, 0)/available.length : null});
  }
  const used = pillars.filter(p => p.score != null);
  const weight = used.reduce((s,p) => s+p.weight,0);
  const composite = weight ? used.reduce((s,p) => s+p.score*p.weight,0)/weight : null;
  let verdict = composite == null ? 'Insufficient data' : composite >= config.verdict.invest ? 'Invest' :
    composite >= config.verdict.watchlist ? 'Watchlist' : 'Avoid';
  const caps = [result.metrics['Beneish M-Score'] != null && result.metrics['Beneish M-Score'] > config.caps['Beneish M-Score'],
    result.metrics["Altman Z''"] != null && result.metrics["Altman Z''"] < config.caps["Altman Z''"],
    result.metrics['CFO / PAT'] != null && result.metrics['CFO / PAT'] < config.caps['CFO / PAT']].filter(Boolean).length;
  if (caps >= 2) verdict = 'Avoid';
  else if (caps && verdict === 'Invest') verdict = 'Watchlist';
  const coverage = Math.round(100 * pillars.reduce((sum, pillar) =>
    sum + pillar.weight * pillar.coverage / 100, 0));
  return {pillars, composite, verdict, caps, coverage};
}

const escapeHtml = text => String(text).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function render(parsed, result, scores, company) {
  const {pillars, composite, verdict, coverage} = scores;
  const bars = pillars.map(p => `<div class="pillar"><div><strong>${escapeHtml(p.name)}</strong><small>Coverage ${p.coverage}%</small></div><div class="track"><span style="width:${p.score ?? 0}%"></span></div><b>${p.score == null ? '—' : Math.round(p.score)}</b></div>`).join('');
  const checks = pillars.flatMap(p => p.items.map(item => `<tr><td>${escapeHtml(item.metric)}</td><td>${item.value == null ? '—' : Number(item.value.toFixed(3))}</td><td>${item.points == null ? 'Skipped' : Math.round(item.points)+'/100'}</td></tr>`)).join('');
  const sales = parsed.data.sales, pat = parsed.data.pat;
  const trends = parsed.years.map((year, i) => `<tr><td>${escapeHtml(year)}</td><td>${sales[i] ?? '—'}</td><td>${pat[i] ?? '—'}</td></tr>`).join('');
  $('analysis').innerHTML = `<div class="analysis-head"><h2>${escapeHtml(company)} — IPO Analysis</h2><span>${escapeHtml(parsed.years[0])}–${escapeHtml(parsed.years.at(-1))} · ${parsed.years.length} years</span></div>
    <div class="score-grid"><div class="verdict ${verdict.toLowerCase().replace(' ','-')}"><small>Verdict</small><strong>${escapeHtml(verdict)}</strong><span>${composite == null ? '—' : Math.round(composite)} / 100 composite score · ${coverage}% data coverage</span></div><div class="pillars">${bars}</div></div>
    <div class="detail-grid"><section><h3>Key Financial Trends</h3><table><thead><tr><th>Year</th><th>Sales</th><th>Net profit</th></tr></thead><tbody>${trends}</tbody></table></section>
    <section><h3>Forensic & Quality Checks</h3><table><thead><tr><th>Metric</th><th>Value</th><th>Score</th></tr></thead><tbody>${checks}</tbody></table>${result.lossMaking ? '<p class="warning">Loss making year found. CFO / PAT excludes years with zero or negative PAT.</p>' : ''}</section></div>`;
  $('analysis').hidden = false;
  $('analysis').scrollIntoView({behavior:'smooth', block:'start'});
}

const configPromise = fetch('analyzer_config.json').then(response => {
  if (!response.ok) throw new Error('Analysis configuration could not be loaded.');
  return response.json();
});
$('upload').addEventListener('change', async event => {
  const file = event.target.files[0];
  if (!file) return;
  $('upload-status').textContent = `Analyzing ${file.name}…`;
  try {
    if (typeof XLSX === 'undefined') throw new Error('Excel reader could not load. Check your connection and try again.');
    const workbook = XLSX.read(await file.arrayBuffer(), {type:'array', cellDates:true});
    const parsed = parseWorkbook(workbook);
    const result = calculate(parsed);
    const scores = score(result, await configPromise);
    render(parsed, result, scores, $('company').value || file.name.replace(/\.xlsx?$/i,''));
    $('upload-status').textContent = `${file.name} analyzed in this browser.`;
  } catch (error) {
    $('upload-status').textContent = error.message;
  }
  event.target.value = '';
});
