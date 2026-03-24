import { Client, Storage } from 'node-appwrite';
import ExcelJS from 'exceljs';

const APPWRITE_ENDPOINT = 'https://cloud.appwrite.io/v1';
const APPWRITE_PROJECT  = process.env.APPWRITE_PROJECT_ID;
const APPWRITE_API_KEY  = process.env.APPWRITE_API_KEY;
const STORAGE_BUCKET_ID = process.env.STORAGE_BUCKET_ID;
const EXCEL_FILE_ID     = process.env.EXCEL_FILE_ID;
const GROQ_API_KEY      = process.env.GROQ_API_KEY;
const GROQ_MODEL        = 'llama-3.3-70b-versatile';
const GROQ_ENDPOINT     = 'https://api.groq.com/openai/v1/chat/completions';
const DESC_JP_MAX       = 1000;
const DESC_EN_MAX       = 1000;
const OVERWRITE         = process.env.OVERWRITE_EXISTING === 'true';
const DELAY_MS          = 2500;

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function callGroq(prompt, log) {
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const res = await fetch(GROQ_ENDPOINT, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${GROQ_API_KEY}`,
          'Content-Type': 'application/json',
          'User-Agent': 'GoldenGaiHopper/1.0',
        },
        body: JSON.stringify({
          model: GROQ_MODEL,
          messages: [{ role: 'user', content: prompt }],
          max_tokens: 600,
          temperature: 0.72,
        }),
      });
      if (res.status === 429) {
        const wait = 30000 * (attempt + 1);
        log(`Rate limited — waiting ${wait/1000}s...`);
        await sleep(wait);
        continue;
      }
      if (!res.ok) { log(`Groq HTTP ${res.status}`); return ''; }
      const data = await res.json();
      return data.choices?.[0]?.message?.content?.trim() || '';
    } catch (e) {
      log(`Groq error (attempt ${attempt+1}): ${e.message}`);
      if (attempt < 2) await sleep(8000);
    }
  }
  return '';
}

async function generateDescriptions(nameJp, nameEn, content, log) {
  const src = content.slice(0, 3000);
  const promptJp = `あなたは新宿ゴールデン街の専門ガイド編集者です。以下の「${nameJp}」に関する情報を元に、Google検索で上位表示されやすい日本語の紹介文を書いてください。【条件】1000文字以内（厳守）、です・ます調、バーの雰囲気・特徴・おすすめポイントを含める、「新宿ゴールデン街」「${nameJp}」を自然に含める、最初の一文でバーの個性を伝える。【参考情報】${src} 紹介文のみを出力:`;
  const promptEn = `You are an expert editor for a Golden Gai bar guide in Shinjuku, Tokyo. Write a compelling SEO-friendly English description for "${nameEn || nameJp}". Max 1000 characters, warm engaging tone for tourists, mention atmosphere and character, naturally include "Golden Gai", "Shinjuku", "${nameEn || nameJp}". Open with a vivid first sentence. Source: ${src} Write only the description:`;

  log(`  Generating JP...`);
  const descJp = (await callGroq(promptJp, log)).slice(0, DESC_JP_MAX);
  await sleep(DELAY_MS);
  log(`  Generating EN...`);
  const descEn = (await callGroq(promptEn, log)).slice(0, DESC_EN_MAX);
  await sleep(DELAY_MS);
  return { descJp, descEn };
}

export default async ({ req, res, log, error }) => {
  log('gg-generate-descriptions starting...');
  log(`Mode: ${OVERWRITE ? 'OVERWRITE all' : 'skip existing'}`);

  const client = new Client()
    .setEndpoint(APPWRITE_ENDPOINT)
    .setProject(APPWRITE_PROJECT)
    .setKey(APPWRITE_API_KEY);
  const storage = new Storage(client);

  log('Downloading Excel...');
  let fileBuffer;
  try {
    const fileData = await storage.getFileDownload(STORAGE_BUCKET_ID, EXCEL_FILE_ID);
    fileBuffer = Buffer.from(fileData);
  } catch (e) {
    error(`Failed to download Excel: ${e.message}`);
    return res.json({ success: false, error: e.message }, 500);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileBuffer);
  const ws = workbook.getWorksheet('BarRemarks') || workbook.worksheets[0];

  let generated = 0, skipped = 0, errors = 0;

  for (let rowNum = 3; rowNum <= ws.rowCount; rowNum++) {
    const row     = ws.getRow(rowNum);
    const uuid    = String(row.getCell(1).value || '').trim();
    const nameJp  = String(row.getCell(2).value || '').trim();
    const nameEn  = String(row.getCell(3).value || '').trim();
    const content = String(row.getCell(4).value || '').trim();
    const existJp = String(row.getCell(7).value || '').trim();

    if (!uuid || uuid.startsWith('Fill')) continue;
    if (!nameJp || ['*', '?'].includes(nameJp)) continue;
    if (!content) { skipped++; continue; }
    if (existJp && !OVERWRITE) { skipped++; continue; }

    log(`\n🍺 [${rowNum-2}] ${nameJp}`);
    const { descJp, descEn } = await generateDescriptions(nameJp, nameEn, content, log);

    if (!descJp && !descEn) { errors++; continue; }

    row.getCell(7).value = descJp;
    row.getCell(8).value = descEn;
    row.commit();
    generated++;
    log(`  Done — JP: ${descJp.length}c / EN: ${descEn.length}c`);
    await sleep(1000);
  }

  log(`Uploading updated Excel... (${generated} generated)`);
  try {
    const updatedBuffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([updatedBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    await storage.deleteFile(STORAGE_BUCKET_ID, EXCEL_FILE_ID);
    await storage.createFile(STORAGE_BUCKET_ID, EXCEL_FILE_ID, blob, ['read(\"any\")']);
    log('Excel uploaded successfully');
  } catch (e) { error(`Failed to upload Excel: ${e.message}`); }

  return res.json({ success: true, generated, skipped, errors });
};
