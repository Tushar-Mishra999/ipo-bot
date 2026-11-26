import http from 'node:http';
import {writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {GoogleGenAI} from '@google/genai';

// Cloud Run uses its attached service account. This supports local development.
if (process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON && !process.env.GOOGLE_APPLICATION_CREDENTIALS) {
  const path = join(tmpdir(), 'ipo-gcp-service-account.json');
  writeFileSync(path, process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON, {mode: 0o600});
  process.env.GOOGLE_APPLICATION_CREDENTIALS = path;
}

const origin = process.env.IPO_DASHBOARD_ORIGIN || 'https://tushar-mishra999.github.io';
const project = process.env.GOOGLE_CLOUD_PROJECT;
const location = process.env.GOOGLE_CLOUD_LOCATION || 'us-central1';
const ai = new GoogleGenAI({vertexai: true, project, location});

const factorMeaning = {
  'Earnings Quality': 'Shows whether reported profit is backed by cash from operations. A weak score can mean profits are slow to turn into cash.',
  'Forensic Risk': 'Looks for unusual accounting patterns, such as rising receivables or reliance on other income. A weak score is a reason to investigate, not proof of wrongdoing.',
  'Financial Strength': 'Shows how comfortably the company can handle debt and interest payments. A weak score means less room for setbacks.',
  'Profitability and Growth': 'Shows whether sales, profits, returns, and operating margins are improving. A strong score suggests healthy business momentum.',
  'Cash and Capital Allocation': 'Shows whether operations generate cash after investment and whether cash supports dividends and debt. A weak score means less financial flexibility.'
};

const verdictAction = {
  Invest: 'What this means: The financials support a closer look, but this score alone is not a reason to subscribe. Check the IPO price, prospectus, and risks before deciding.',
  Watchlist: 'What this means: Wait rather than subscribe based on this score. Reassess after listing when the price and more financial evidence are available.',
  Avoid: 'What this means: The current financial evidence does not support subscribing. Reconsider only if the weak factors improve.',
  'Insufficient data': 'What this means: Wait for fuller financial statements before deciding whether to subscribe or buy after listing.'
};

function send(response, status, body) {
  response.writeHead(status, {
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json; charset=utf-8'
  });
  response.end(JSON.stringify(body));
}

async function explain(payload) {
  const verdict = payload.verdict;
  if (!['Invest', 'Watchlist', 'Avoid', 'Insufficient data'].includes(verdict)) {
    throw new Error('Invalid verdict');
  }
  const facts = {
    company: String(payload.company || 'the company').slice(0, 120),
    verdict,
    composite: payload.composite,
    coverage: payload.coverage,
    caps: payload.caps,
    pillars: Array.isArray(payload.pillars) ? payload.pillars.slice(0, 5).map(pillar => ({
      name: pillar.name,
      score: pillar.score,
      coverage: pillar.coverage,
      meaning: factorMeaning[pillar.name] || ''
    })) : [],
    metrics: payload.metrics || {}
  };
  const prompt = `Write 55 to 75 words explaining the reasons behind this fixed IPO verdict for a first-time investor. Explain what the strongest and weakest measured factors usually mean for a business, using the factor meanings supplied. Discuss any verdict caps and material data gaps. Focus on interpretation, not a list of scores; use at most one number. Do not state a verdict or tell the reader what to buy, sell, or subscribe to. Do not imply that a forensic warning proves wrongdoing. Use only the supplied facts.\n\n${JSON.stringify(facts)}`;
  const response = await ai.models.generateContent({
    model: 'gemini-2.5-flash',
    contents: prompt,
    config: {
      maxOutputTokens: 1500,
      temperature: 0.2,
      thinkingConfig: {thinkingBudget: 0},
      systemInstruction: 'Explain a fixed financial score in simple everyday terms. The verdict and action will be added separately. Return one plain-text paragraph only.'
    }
  });
  const text = response.text?.trim();
  if (!text) throw new Error('Empty AI response');
  return `${text}\n\n${verdictAction[verdict]}`;
}

http.createServer((request, response) => {
  if (request.method === 'OPTIONS') return send(response, 204, {});
  if (request.method === 'GET' && request.url === '/health') return send(response, 200, {ok: true});
  if (request.method !== 'POST' || request.url !== '/explain' || request.headers.origin !== origin) {
    return send(response, 403, {error: 'Forbidden'});
  }
  const chunks = [];
  let size = 0;
  request.on('data', chunk => {
    size += chunk.length;
    if (size > 20000) request.destroy();
    else chunks.push(chunk);
  });
  request.on('end', async () => {
    try {
      const payload = JSON.parse(Buffer.concat(chunks).toString('utf8'));
      const explanation = await explain(payload);
      send(response, 200, {explanation});
    } catch (error) {
      send(response, 400, {error: error.message});
    }
  });
}).listen(Number(process.env.PORT || 8080));
