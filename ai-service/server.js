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
    pillars: Array.isArray(payload.pillars) ? payload.pillars.slice(0, 5) : [],
    metrics: payload.metrics || {}
  };
  const prompt = `Explain this IPO scoring result in plain language in 90 to 110 words. The supplied verdict is final and must be stated exactly. Explain the strongest and weakest measured factors and any verdict caps. Mention limited data coverage when material. Use only the supplied facts. Do not invent valuations, transactions, or recommendations.\n\n${JSON.stringify(facts)}`;
  const response = await ai.models.generateContent({
    model: 'gemini-2.5-flash',
    contents: prompt,
    config: {
      maxOutputTokens: 1500,
      temperature: 0.2,
      thinkingConfig: {thinkingBudget: 0},
      systemInstruction: 'Explain a fixed financial score in simple terms. Preserve the verdict exactly. Return plain text only.'
    }
  });
  const text = response.text?.trim();
  if (!text) throw new Error('Empty AI response');
  const verdictLabels = text.match(/\b(?:Invest|Watchlist|Avoid)\b/gi) || [];
  if (verdictLabels.some(label => label.toLowerCase() !== verdict.toLowerCase())) {
    throw new Error('AI response conflicted with the calculated verdict');
  }
  return verdictLabels.length ? text : `The verdict is ${verdict}. ${text}`;
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
