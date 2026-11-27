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
const label = value => String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ').replace(/[^a-z0-9]/g, '');
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
  const headers = rows.filter(row => label(row[0]) === 'reportdate');
  const header = headers.reduce((best, row) => row.slice(1).filter(v => v != null && String(v).trim() && !/ttm|quarter/i.test(String(v))).length >
    (best?.slice(1).filter(v => v != null && String(v).trim() && !/ttm|quarter/i.test(String(v))).length ?? 0) ? row : best, null);
  if (!header) throw new Error('Report Date row missing from Data Sheet.');
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
  const indexed = new Map();
  for (const row of rows) {
    const key = label(row[0]);
    if (key && !indexed.has(key)) indexed.set(key, row);
    else if (key && indexed.get(key) && row.slice(1).some(v => number(v) != null) &&
      !indexed.get(key).slice(1).some(v => number(v) != null)) indexed.set(key, row);
  }
  const data = {}, missingLabels = [], blankCells = {};
  for (const [key, aliases] of Object.entries(rowNames)) {
    const row = aliases.map(name => indexed.get(label(name))).find(Boolean);
    if (!row) missingLabels.push(aliases[0]);
    blankCells[key] = columns.map(i => !!row && number(row[i]) == null);
    data[key] = columns.map(i => row ? (number(row[i]) ?? 0) : null);
  }
  return {years, data, missingLabels, blankCells};
}

function calculate(parsed) {
  const {years, data} = parsed;
  const n = years.length;
  const diagnostics = {};
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
  for (const [name, key] of [['Sales CAGR','sales'], ['PAT CAGR','pat']]) {
    if (n === 2) diagnostics[name] = {notes: ['1-year growth: this trend rests on only two data points.']};
    if (n < 2) diagnostics[name] = {skipReason: `Needs at least 2 years (found ${n})`};
    else if (metrics[name] == null) diagnostics[name] = {skipReason: `${rowNames[key][0]} must be positive in the first and last years`};
  }
  metrics['ROE'] = divide(latest('pat'), equity(n - 1));
  metrics['EBITDA margin trend'] = n > 1 ? difference(margin(n - 1), margin(0)) : null;
  const capex = n > 1 ? add(difference(latest('netBlock'), previous('netBlock')),
    difference(latest('cwip'), previous('cwip')), latest('depreciation')) : null;
  const fcf = difference(latest('cfo'), capex);
  metrics['FCF / PAT'] = latest('pat') > 0 ? divide(fcf, latest('pat')) : null;
  metrics['Dividend payout vs cash'] = latest('cfo') > 0 && latest('dividend') != null ?
    divide(latest('dividend'), latest('cfo')) : null;
  metrics['Cash / Borrowings'] = divide(latest('cash'), latest('borrowings'));
  const beneish = calculateBeneish(parsed);
  metrics['Beneish M-Score'] = beneish.value;
  diagnostics['Beneish M-Score'] = beneish.diagnostic;
  const needs = {
    'CFO / PAT': ['pat','cfo'], Accruals: ['pat','cfo','assets'], 'Receivable days change': ['receivables','sales'],
    'Other income share': ['otherIncome','pbt'], "Altman Z''": ['receivables','inventory','cash','liabilities','assets','reserves','pbt','interest','equity','borrowings'],
    'Interest cover': ['pbt','interest'], 'Debt / Equity': ['borrowings','equity','reserves'],
    ROE: ['pat','equity','reserves'], 'EBITDA margin trend': ['pbt','interest','depreciation','otherIncome','sales'],
    'FCF / PAT': ['cfo','netBlock','cwip','depreciation','pat'], 'Dividend payout vs cash': ['dividend','cfo'],
    'Cash / Borrowings': ['cash','borrowings']
  };
  const twoYears = new Set(['Receivable days change','EBITDA margin trend','FCF / PAT']);
  for (const [name, value] of Object.entries(metrics)) if (value == null && !diagnostics[name]) {
    const missingKey = needs[name]?.find(key => data[key]?.[n-1] == null);
    const reason = twoYears.has(name) && n < 2 ? `Needs at least 2 years (found ${n})` :
      missingKey ? `${rowNames[missingKey][0]} row missing` :
      name === 'CFO / PAT' ? 'No year has positive net profit and operating cash data' :
      name === 'FCF / PAT' ? 'Net profit must be positive and required capex data available' :
      name === 'Dividend payout vs cash' ? 'Cash from Operating Activity must be positive' :
      'A required denominator is zero';
    diagnostics[name] = {skipReason: reason};
  }
  return {metrics, diagnostics, extra: {receivableDays: daysNow, EBITDA: margin(n-1), capex, fcf}, lossMaking: data.pat.some(v => v != null && v <= 0)};
}

function calculateBeneish({years, data, blankCells = {}}) {
  const diagnostic = {notes: []};
  if (years.length < 2) return {value: null, diagnostic: {skipReason: `Needs at least 2 years (found ${years.length})`}};
  const core = ['sales','receivables','assets','netBlock','pat','cfo'];
  const i = years.length - 1, j = i - 1;
  for (const key of core) if (!data[key] || data[key][i] == null || data[key][j] == null)
    return {value: null, diagnostic: {skipReason: `${rowNames[key][0]} row missing or has non-numeric data in the last two years`}};
  const optional = ['raw','inventoryChange','power','manufacturing','selling','depreciation','borrowings','liabilities','inventory','cash'];
  const get = (key, index) => {
    const name = rowNames[key][0];
    if (data[key]?.[index] != null) {
      if (blankCells[key]?.[index] && !diagnostic.notes.includes(`${name} blank in ${years[index]}; estimated as 0.`))
        diagnostic.notes.push(`${name} blank in ${years[index]}; estimated as 0.`);
      return data[key][index];
    }
    if (!diagnostic.notes.includes(`${name} row missing; estimated as 0.`)) diagnostic.notes.push(`${name} row missing; estimated as 0.`);
    return 0;
  };
  for (const key of optional) { get(key, i); get(key, j); }
  const ratio = (a, b, name, reason) => {
    if (!Number.isFinite(a) || !Number.isFinite(b) || b === 0) {
      diagnostic.notes.push(`${name} set to 1: ${reason || 'its denominator is zero'}.`);
      return 1;
    }
    const value = a / b;
    if (value > 3) {
      diagnostic.notes.push(`${name} capped at 3: ${reason || 'the denominator was very small, making the ratio extreme'}.`);
      return 3;
    }
    return value;
  };
  const intermediate = (a, b) => b === 0 ? 0 : a / b;
  const sales = k => get('sales', k), assets = k => get('assets', k);
  const gross = k => 1 - intermediate(get('raw',k)-get('inventoryChange',k)+get('power',k)+get('manufacturing',k),sales(k));
  const current = k => get('receivables',k)+get('inventory',k)+get('cash',k);
  const quality = k => 1 - intermediate(get('netBlock',k)+current(k),assets(k));
  const depRate = k => intermediate(get('depreciation',k),get('depreciation',k)+get('netBlock',k));
  const leverage = k => intermediate(get('borrowings',k)+get('liabilities',k),assets(k));
  const dsri = ratio(intermediate(get('receivables',i),sales(i)),intermediate(get('receivables',j),sales(j)),'DSRI','prior-year receivables were very small, so the ratio is extreme');
  const gmi = ratio(gross(j),gross(i),'GMI','current gross margin was zero or very small');
  const aqi = ratio(quality(i),quality(j),'AQI','prior-year asset quality was zero or very small');
  const sgi = ratio(sales(i),sales(j),'SGI','prior-year sales were zero or very small');
  const depi = ratio(depRate(j),depRate(i),'DEPI','current depreciation rate was zero or very small');
  const sgai = ratio(intermediate(get('selling',i),sales(i)),intermediate(get('selling',j),sales(j)),'SGAI','prior-year SG&A was zero or very small');
  const lvgi = ratio(leverage(i),leverage(j),'LVGI','prior-year leverage was zero or very small');
  const tata = ratio(get('pat',i)-get('cfo',i),assets(i),'TATA','current assets were zero');
  const value = -4.84 + .92*dsri + .528*gmi + .404*aqi + .892*sgi + .115*depi - .172*sgai + 4.679*tata - .327*lvgi;
  return {value, diagnostic};
}

function preferenceWeights(config, preferences) {
  const settings = config.preferences;
  const raw = Object.fromEntries(Object.entries(config.pillars).map(([name, pillar]) => [name, pillar.weight * 100]));
  for (const key of ['holding', 'safety', 'income']) {
    const delta = (preferences[key] - settings.default) / settings.default;
    for (const [name, amount] of Object.entries(settings[key].adjustments)) raw[name] += amount * delta;
  }
  for (const name of Object.keys(raw)) raw[name] = Math.max(settings.weight_floor, raw[name]);
  const total = Object.values(raw).reduce((sum, value) => sum + value, 0);
  return Object.fromEntries(Object.entries(raw).map(([name, value]) => [name, value / total]));
}
function score(result, config, preferences = Object.fromEntries(['holding', 'safety', 'income'].map(key => [key, config.preferences.default]))) {
  const weights = preferenceWeights(config, preferences);
  const pillars = [];
  for (const [name, definition] of Object.entries(config.pillars)) {
    const items = Object.entries(definition.metrics).map(([metric, [bad, good]]) => {
      const value = result.metrics[metric];
      const points = value == null ? null : Math.max(0, Math.min(100, 100*(value-bad)/(good-bad)));
      return {metric, value, points, diagnostic: result.diagnostics?.[metric]};
    });
    const available = items.filter(item => item.points != null);
    pillars.push({name, weight: weights[name], items,
      coverage: Math.round(100*available.length/items.length),
      score: available.length ? available.reduce((s, item) => s+item.points, 0)/available.length : null});
  }
  const used = pillars.filter(p => p.score != null);
  const weight = used.reduce((s,p) => s+p.weight,0);
  const composite = weight ? used.reduce((s,p) => s+p.score*p.weight,0)/weight : null;
  let verdict = composite == null ? 'Insufficient data' : composite >= config.verdict.invest ? 'Invest' :
    composite >= config.verdict.watchlist ? 'Watchlist' : 'Avoid';
  const baseCaps = [
    result.metrics['Beneish M-Score'] != null && result.metrics['Beneish M-Score'] > config.caps['Beneish M-Score'] ? 'Beneish M-Score warning' : null,
    result.metrics["Altman Z''"] != null && result.metrics["Altman Z''"] < config.caps["Altman Z''"] ? 'Altman Z distress warning' : null,
    result.metrics['CFO / PAT'] != null && result.metrics['CFO / PAT'] < config.caps['CFO / PAT'] ? 'Weak profit-to-cash conversion' : null
  ].filter(Boolean);
  const safetyRule = config.preferences.safety_cap;
  const safetyCaps = preferences.safety >= safetyRule.min_slider ? [
    result.metrics['Debt / Equity'] != null && result.metrics['Debt / Equity'] > safetyRule.debt_equity_above ? 'Debt/Equity exceeds safety limit' : null,
    result.metrics['Interest cover'] != null && result.metrics['Interest cover'] < safetyRule.interest_cover_below ? 'Interest cover is below safety limit' : null
  ].filter(Boolean) : [];
  if (baseCaps.length >= 2) verdict = 'Avoid';
  else if ((baseCaps.length || safetyCaps.length) && verdict === 'Invest') verdict = 'Watchlist';
  const coverage = Math.round(100 * pillars.reduce((sum, pillar) =>
    sum + pillar.weight * pillar.coverage / 100, 0));
  return {pillars, composite, verdict, caps: baseCaps.length, capReasons: [...baseCaps, ...safetyCaps], coverage, weights};
}

const escapeHtml = text => String(text).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const formatMetric = (value, format) => {
  if (value == null) return '—';
  if (format === 'percent') return `${(value * 100).toFixed(1).replace(/\.0$/, '')}%`;
  if (format === 'points') return `${(value * 100).toFixed(1).replace(/\.0$/, '')} pp`;
  if (format === 'multiple') return `${Number(value.toFixed(2))}x`;
  return String(Number(value.toFixed(format === 'decimal' ? 2 : 3)));
};
let infoSequence = 0;
function infoIcon(title, lines) {
  const id = `metric-tip-${++infoSequence}`;
  const content = lines.filter(Boolean).map(line => `<span>${escapeHtml(line)}</span>`).join('');
  return `<span class="info-wrap"><button type="button" class="info-icon" aria-label="About ${escapeHtml(title)}" aria-describedby="${id}">i</button><span class="info-tip" id="${id}" role="tooltip">${content}</span></span>`;
}
function metricTip(name, info, config, diagnostic) {
  const text = info.metrics[name];
  const thresholds = Object.values(config.pillars).map(p => p.metrics[name]).find(Boolean);
  if (!text || !thresholds) return '';
  const [bad, good] = thresholds;
  const value = number => formatMetric(name === 'EBITDA margin trend' && number < 0 ? -number : number, text.format);
  return infoIcon(name, [text.description, `Good: ${text.good.replace('{value}', value(good))}`, `Watch: ${text.watch.replace('{value}', value(bad))}`, ...(diagnostic?.notes || []), diagnostic?.skipReason && `Skipped: ${diagnostic.skipReason}`]);
}
function trendChart(years, sales, profit) {
  const width = 620, height = 285, left = 58, right = 58, top = 30, bottom = 45;
  const plotWidth = width-left-right, plotHeight = height-top-bottom;
  const salesValues = sales.filter(v => v != null), profitValues = profit.filter(v => v != null);
  if (!salesValues.length && !profitValues.length) return '<p>No sales or profit data available.</p>';
  const salesMax = Math.max(1, ...salesValues) * 1.12;
  const profitMax = Math.max(1, ...profitValues) * 1.12;
  const profitMin = Math.min(0, ...profitValues);
  const salesY = value => top + (salesMax-value)/salesMax*plotHeight;
  const profitY = value => top + (profitMax-value)/(profitMax-profitMin)*plotHeight;
  const x = i => left + (i+0.5)*plotWidth/years.length;
  const baseline = salesY(0);
  const ticks = Array.from({length: 5}, (_,i) => i/4).map(fraction => {
    const y = top + (1-fraction)*plotHeight;
    const salesTick = salesMax*fraction;
    const profitTick = profitMin+(profitMax-profitMin)*fraction;
    return `<line x1="${left}" x2="${width-right}" y1="${y}" y2="${y}" stroke="#364064" stroke-dasharray="4 5"/><text x="${left-8}" y="${y+4}" text-anchor="end" fill="#ba96ff" font-size="11">${Math.round(salesTick).toLocaleString()}</text><text x="${width-right+8}" y="${y+4}" fill="#72f6a1" font-size="11">${Math.round(profitTick).toLocaleString()}</text>`;
  }).join('');
  const bars = sales.map((value,i) => value == null ? '' :
    `<rect x="${x(i)-Math.min(25,plotWidth/years.length*0.34)}" y="${Math.min(salesY(value),baseline)}" width="${Math.min(50,plotWidth/years.length*0.68)}" height="${Math.max(1,Math.abs(baseline-salesY(value)))}" rx="5" fill="url(#salesGradient)"><title>${escapeHtml(years[i])} sales: ${value.toLocaleString()}</title></rect>`).join('');
  const segments = [];
  for (let i=1;i<profit.length;i++) if (profit[i-1] != null && profit[i] != null)
    segments.push(`<line x1="${x(i-1)}" y1="${profitY(profit[i-1])}" x2="${x(i)}" y2="${profitY(profit[i])}" stroke="#72f6a1" stroke-width="3"/>`);
  const points = profit.map((value,i) => value == null ? '' :
    `<circle cx="${x(i)}" cy="${profitY(value)}" r="5" fill="#72f6a1" stroke="#17213d" stroke-width="2"><title>${escapeHtml(years[i])} profit: ${value.toLocaleString()}</title></circle>`).join('');
  const labels = years.map((year,i) => `<text x="${x(i)}" y="${height-12}" text-anchor="middle" fill="#c6c9e1" font-size="11">${escapeHtml(year)}</text>`).join('');
  return `<div class="chart-legend"><span class="sales-dot"></span>Sales · left axis <span class="profit-dot"></span>Net profit · right axis</div><svg class="trend-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="Sales bars on the left axis and net profit line on the right axis by year"><defs><linearGradient id="salesGradient" x1="0" y1="0" x2="0" y2="1"><stop stop-color="#aa71ff"/><stop offset="1" stop-color="#6d48d3"/></linearGradient></defs>${ticks}${bars}${segments.join('')}${points}${labels}</svg>`;
}
function cashflowChart(years, cashflow, profit) {
  cashflow = cashflow || [];
  if (!cashflow.some(value => value != null)) return '<p>Operating cash flow is unavailable in this workbook.</p>';
  const width = 620, height = 265, left = 58, right = 20, top = 24, bottom = 42;
  const plotWidth = width-left-right, plotHeight = height-top-bottom;
  const values = [...cashflow, ...profit].filter(value => value != null);
  const low = Math.min(0, ...values), high = Math.max(0, ...values);
  const padding = Math.max(1, (high-low)*0.12);
  const min = low < 0 ? low-padding : 0, max = high > 0 ? high+padding : padding;
  const y = value => top+(max-value)/(max-min)*plotHeight;
  const zero = y(0);
  const x = index => left+(index+0.5)*plotWidth/years.length;
  const barWidth = Math.min(18, plotWidth/years.length*0.24);
  const ticks = Array.from({length:5}, (_,index) => min+(max-min)*index/4).map(value =>
    `<line x1="${left}" x2="${width-right}" y1="${y(value)}" y2="${y(value)}" stroke="${Math.abs(value)<0.0001?'#8792b2':'#364064'}" stroke-dasharray="4 5"/><text x="${left-8}" y="${y(value)+4}" text-anchor="end" fill="#bfc9e4" font-size="11">${Math.round(value).toLocaleString()}</text>`).join('');
  const bars = years.map((year,index) => [[cashflow[index], '#72f6a1', -barWidth-2, 'Operating cash flow'], [profit[index], '#a578f0', 2, 'Net profit']].map(([value,color,offset,label]) => value == null ? '' :
    `<rect x="${x(index)+offset}" y="${Math.min(y(value),zero)}" width="${barWidth}" height="${Math.max(1,Math.abs(zero-y(value)))}" rx="3" fill="${color}"><title>${escapeHtml(year)} ${label}: ${value.toLocaleString()}</title></rect>`).join('')).join('');
  const labels = years.map((year,index) => `<text x="${x(index)}" y="${height-10}" text-anchor="middle" fill="#c6c9e1" font-size="11">${escapeHtml(year)}</text>`).join('');
  return `<div class="cashflow-legend"><span style="background:#72f6a1"></span>Operating cash flow <span style="background:#a578f0;margin-left:10px"></span>Net profit</div><svg class="cashflow-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="Grouped bars comparing operating cash flow and net profit by year">${ticks}<line x1="${left}" x2="${width-right}" y1="${zero}" y2="${zero}" stroke="#8792b2"/>${bars}${labels}</svg>`;
}
function scoreColor(points) {
  if (points == null) return 'muted';
  if (points >= 75) return 'good';
  if (points >= 40) return 'caution';
  return 'bad';
}
const AI_API = window.IPO_AI_ENDPOINT || '';
async function loadExplanation(company, result, scores, preferences) {
  if (!AI_API) {
    return 'AI explanation is awaiting server setup.';
  }
  try {
    const response = await fetch(`${AI_API}/explain`, {
      method: 'POST', headers: {'Content-Type':'application/json'},
      body: JSON.stringify({company, verdict:scores.verdict, composite:scores.composite,
        coverage:scores.coverage, caps:scores.caps, capReasons:scores.capReasons, preferences,
        pillars:scores.pillars.map(({name,score,coverage}) => ({name,score,coverage})),
        metrics:result.metrics})
    });
    if (!response.ok) throw new Error('AI analysis is temporarily unavailable.');
    const body = await response.json();
    return body.explanation;
  } catch (error) {
    return error.message;
  }
}
const preferences = {};
let currentAnalysis = null;
let explanationVersion = 0;
let explanationTimer = null;
const preferenceSummary = () => [
  preferences.holding <= 2 ? 'short holding horizon' : preferences.holding >= 8 ? 'long holding horizon' : 'medium holding horizon',
  preferences.safety >= 7 ? 'safety first' : preferences.safety <= 3 ? 'comfortable with risk' : 'balanced safety',
  preferences.income >= 7 ? 'income important' : preferences.income <= 3 ? 'little income need' : 'moderate income need'
].join(' · ');
function ipoContext(company) {
  const clean = value => value.toLowerCase().replace(/[^a-z0-9]/g, '').replace(/(?:industries|limited|ltd)$/,'');
  const name = clean(company);
  const match = (window.IPO_CONTEXT || []).find(item => {
    const candidate = clean(item.name);
    return name.length > 5 && (candidate === name || candidate.includes(name) || name.includes(candidate));
  });
  if (!match) return '<p>GMP and subscription status are unavailable for this uploaded file.</p>';
  const today = new Intl.DateTimeFormat('en-CA', {timeZone:'Asia/Kolkata', year:'numeric', month:'2-digit', day:'2-digit'}).format(new Date());
  const status = match.open_date && today < match.open_date ? 'Upcoming' : match.close_date && today > match.close_date ? 'Closed' : 'Open';
  const gmp = match.gmp_pct == null ? 'Unavailable' : `${match.gmp_pct >= 0 ? '+' : ''}${Number(match.gmp_pct).toFixed(2)}%`;
  return `<p>GMP: <strong>${gmp}</strong> · Subscription window: <strong>${status}</strong></p><p>GMP is an unofficial market indicator, shown as context only and never part of the score. Subscription multiples are unavailable.</p>`;
}
function signalPanels(scores) {
  const items = scores.pillars.flatMap(p => p.items.filter(item => item.points != null).map(item => ({...item, priority: p.weight})));
  const list = (entries, fallback) => entries.slice(0, 3).map(item => `<li>${escapeHtml(item.metric)} (${Math.round(item.points)}/100)</li>`).join('') || `<li>${fallback}</li>`;
  const flags = items.filter(item => item.points < 50).sort((a,b) => b.priority*(100-b.points)-a.priority*(100-a.points));
  const strengths = items.filter(item => item.points >= 70).sort((a,b) => b.priority*b.points-a.priority*a.points);
  return `<div class="strengths-flags"><section><h3>Top red flags for your goals</h3><ul>${list(flags, 'No strong red flags in the measured metrics')}</ul></section><section><h3>Top strengths for your goals</h3><ul>${list(strengths, 'No strong factors in the measured metrics')}</ul></section></div>`;
}
function render(parsed, result, scores, company, explanation, config, info, scroll = true) {
  const {pillars, composite, verdict} = scores;
  const bars = pillars.map(p => `<div class="pillar"><div><strong>${escapeHtml(p.name)}</strong>${infoIcon(p.name, [info.pillars[p.name]?.description, `High score: ${info.pillars[p.name]?.high}`, `Low score: ${info.pillars[p.name]?.low}`])}</div><div class="track"><span style="width:${p.score ?? 0}%"></span></div><b>${p.score == null ? '—' : Math.round(p.score)}</b></div>`).join('');
  const checks = pillars.flatMap(p => p.items.map(item => `<tr><td>${escapeHtml(parsed.years.length === 2 && / CAGR$/.test(item.metric) ? item.metric.replace('CAGR', '1-year growth') : item.metric)} ${metricTip(item.metric, info, config, item.diagnostic)}</td><td>${formatMetric(item.value, info.metrics[item.metric]?.format)}</td><td class="${scoreColor(item.points)}">${item.points == null ? `Skipped: ${escapeHtml(item.diagnostic?.skipReason || 'Required data unavailable')}` : Math.round(item.points)+'/100'}</td></tr>`)).join('');
  const weights = pillars.map(p => `${(p.weight*100).toFixed(1)}%`).join('/');
  const compositeTip = infoIcon('Composite Score', [info.composite.description.replace('{weights}', weights), `Good: ${info.composite.good.replace('{value}', config.verdict.invest)}`, `Watch: ${info.composite.watch.replace('{value}', config.verdict.watchlist)}`]);
  const weightsUsed = pillars.map(p => `<span>${escapeHtml(p.name)} ${(p.weight*100).toFixed(1)}%</span>`).join('');
  const shortContext = preferences.holding <= config.preferences.short_term.max_slider ? `<div class="short-context"><p class="short-notice">${escapeHtml(config.preferences.short_term.notice)}</p>${ipoContext(company)}</div>` : '';
  const sales = parsed.data.sales, pat = parsed.data.pat;
  const trends = trendChart(parsed.years, sales, pat);
  const cashflowTrend = cashflowChart(parsed.years, parsed.data.cfo, pat);
  $('analysis').innerHTML = `<div class="analysis-head"><h2>${escapeHtml(company)} — IPO Analysis</h2><span>${escapeHtml(parsed.years[0])}–${escapeHtml(parsed.years.at(-1))} · ${parsed.years.length} years</span></div>
    <div class="score-grid"><div class="verdict ${verdict.toLowerCase().replace(' ','-')}"><div class="verdict-text"><small>Verdict</small><strong>${escapeHtml(verdict)}</strong><div class="scored-for">Scored for: ${escapeHtml(preferenceSummary())}</div><div class="weight-list" aria-label="Weights used"><b>Weights used:</b>${weightsUsed}</div><div class="education-note">${escapeHtml(config.preferences.education_note)}</div></div><div class="gauge"><svg viewBox="0 0 200 120" aria-hidden="true"><path class="gauge-bg" d="M 20 95 A 80 80 0 0 1 180 95"/><path class="gauge-value" d="M 20 95 A 80 80 0 0 1 180 95" style="stroke-dasharray:${composite == null ? 0 : Math.max(0,Math.min(100,composite))*2.513} 252"/></svg><div class="gauge-score"><b>${composite == null ? '—' : Math.round(composite)}</b><span>/100</span></div><small>Composite Score ${compositeTip}</small></div>${shortContext}</div><section class="ai-card"><h3>AI Analysis</h3><p id="ai-explanation">${escapeHtml(explanation)}</p></section></div>${signalPanels(scores)}<div class="pillars pillar-grid">${bars}</div>
    <div class="detail-grid"><section><h3>Key Financial Trends</h3>${trends}<h4 class="trend-subhead">Cash from operations vs net profit</h4>${cashflowTrend}</section>
    <section><h3>Forensic & Quality Checks</h3><table><thead><tr><th>Metric</th><th>Value</th><th>Score</th></tr></thead><tbody>${checks}</tbody></table>${result.lossMaking ? '<p class="warning">Loss making year found. CFO / PAT excludes years with zero or negative PAT.</p>' : ''}</section></div>`;
  $('analysis').hidden = false;
  if (scroll) $('analysis').scrollIntoView({behavior:'smooth', block:'start'});
}

const configPromise = fetch('analyzer_config.json').then(response => {
  if (!response.ok) throw new Error('Analysis configuration could not be loaded.');
  return response.json();
});
const infoPromise = fetch('metric_info.json').then(response => {
  if (!response.ok) throw new Error('Metric information could not be loaded.');
  return response.json();
});
function refreshForPreferences() {
  const version = ++explanationVersion;
  if (!currentAnalysis) return;
  clearTimeout(explanationTimer);
  const {parsed, result, company, config, info} = currentAnalysis;
  const scores = score(result, config, preferences);
  render(parsed, result, scores, company, 'Updating explanation…', config, info, false);
  explanationTimer = setTimeout(async () => {
    const explanation = await loadExplanation(company, result, scores, {...preferences});
    if (version === explanationVersion) $('ai-explanation').textContent = explanation;
  }, 400);
}
configPromise.then(config => {
  const settings = config.preferences;
  for (const key of ['holding', 'safety', 'income']) preferences[key] = settings.default;
  const controls = ['holding', 'safety', 'income'].map(key => {
    const item = settings[key];
    return `<div class="preference"><div class="preference-head"><label for="pref-${key}">${escapeHtml(item.label)}</label>${infoIcon(item.label, [item.info])}<output id="pref-${key}-value" for="pref-${key}">${settings.default}</output><button type="button" class="reset-preferences" data-reset="${key}">Reset</button></div><input id="pref-${key}" type="range" min="${settings.min}" max="${settings.max}" step="${settings.step}" value="${settings.default}"><div class="preference-ends"><span>${escapeHtml(item.low)}</span><span>${escapeHtml(item.high)}</span></div></div>`;
  }).join('');
  $('preferences').innerHTML = `<h3>Tell us what you are looking for</h3>${controls}`;
  for (const key of ['holding', 'safety', 'income']) {
    const slider = $(`pref-${key}`);
    slider.addEventListener('input', () => {
      preferences[key] = Number(slider.value);
      $(`pref-${key}-value`).textContent = slider.value;
      refreshForPreferences();
    });
  }
  document.querySelectorAll('[data-reset]').forEach(button => button.addEventListener('click', () => {
    const slider = $(`pref-${button.dataset.reset}`);
    slider.value = settings.default;
    slider.dispatchEvent(new Event('input', {bubbles:true}));
  }));
}).catch(error => { $('preferences').textContent = error.message; });
$('upload').addEventListener('change', async event => {
  const file = event.target.files[0];
  if (!file) return;
  currentAnalysis = null;
  explanationVersion++;
  clearTimeout(explanationTimer);
  const loading = document.createElement('div');
  loading.className = 'analysis-loading';
  loading.setAttribute('role', 'status');
  loading.innerHTML = '<span class="spinner" aria-hidden="true"></span> Preparing the full analysis and AI explanation…';
  $('analysis').hidden = false;
  $('analysis').replaceChildren(loading);
  $('upload-status').textContent = `Analyzing ${file.name}…`;
  try {
    await new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)));
    if (typeof XLSX === 'undefined') throw new Error('Excel reader could not load. Check your connection and try again.');
    const workbook = XLSX.read(await file.arrayBuffer(), {type:'array', cellDates:true});
    const parsed = parseWorkbook(workbook);
    const result = calculate(parsed);
    const config = await configPromise;
    const info = await infoPromise;
    const company = file.name.replace(/\.xlsx?$/i,'');
    let version, scores, explanation;
    do {
      version = explanationVersion;
      scores = score(result, config, preferences);
      explanation = await loadExplanation(company, result, scores, {...preferences});
    } while (version !== explanationVersion);
    currentAnalysis = {parsed, result, company, config, info};
    render(parsed, result, scores, company, explanation, config, info);
    $('upload-status').textContent = `${file.name} analyzed in this browser.`;
  } catch (error) {
    $('upload-status').textContent = error.message;
    $('analysis').hidden = true;
  }
  event.target.value = '';
});
