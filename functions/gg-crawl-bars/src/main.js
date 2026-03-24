import { Client, Storage } from 'node-appwrite';
import ExcelJS from 'exceljs';

const APPWRITE_ENDPOINT = 'https://cloud.appwrite.io/v1';
const APPWRITE_PROJECT  = process.env.APPWRITE_PROJECT_ID;
const APPWRITE_API_KEY  = process.env.APPWRITE_API_KEY;
const STORAGE_BUCKET_ID = process.env.STORAGE_BUCKET_ID;
const EXCEL_FILE_ID     = process.env.EXCEL_FILE_ID;
const GOOGLE_API_KEY    = process.env.GOOGLE_API_KEY;
const SEARCH_ENGINE_ID  = process.env.SEARCH_ENGINE_ID || '616f157829e604251';
const SEARCHES_PER_RUN  = 90;
const MAX_CONTENT_CHARS = 10000;
const DELAY_MS          = 1300;

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function searchBar(barName) {
  const query = `ゴールデン街 ${barName}`;
  const url = new URL('https://www.googleapis.com/customsearch/v1');
  url.searchParams.set('key', GOOGLE_API_KEY);
  url.searchParams.set('cx', SEARCH_ENGINE_ID);
  url.searchParams.set('q', query);
  url.searchParams.set('num', '3');
  url.searchParams.set('lr', 'lang_ja');
  try {
    const res = await fetch(url.toString());
    if (!res.ok) return '';
    const data = await res.json();
    const items = data.items || [];
    return items.map(i => `${i.title || ''}: ${i.snippet || ''}`).join(' | ');
  } catch { return ''; }
}

export default async ({ req, res, log, error }) => {
  log('gg-crawl-bars starting...');
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

  const headerRow = ws.getRow(1);
  let nameCol = 2, contentCol = 4;
  headerRow.eachCell((cell, col) => {
    const v = String(cell.value || '').toLowerCase();
    if (v.includes('barname')) nameCol = col;
    if (v.includes('content')) contentCol = col;
  });

  const today = new Date().toISOString().slice(0, 10);
  let searched = 0, updated = 0, skipped = 0;

  for (let rowNum = 3; rowNum <= ws.rowCount; rowNum++) {
    if (searched >= SEARCHES_PER_RUN) { log('Daily limit reached'); break; }
    const row = ws.getRow(rowNum);
    const barName = row.getCell(nameCol).value;
    if (!barName) continue;
    const barNameStr = String(barName).trim();
    if (!barNameStr || ['*', '?'].includes(barNameStr) || barNameStr.startsWith('Fill')) { skipped++; continue; }

    const currentContent = String(row.getCell(contentCol).value || '');
    log(`[${searched + 1}/${SEARCHES_PER_RUN}] Searching: ${barNameStr}`);
    const newContent = await searchBar(barNameStr);

    if (newContent) {
      let combined = currentContent + `\n[${today}] ${newContent}`;
      if (combined.length > MAX_CONTENT_CHARS) combined = combined.slice(-MAX_CONTENT_CHARS);
      row.getCell(contentCol).value = combined;
      row.commit();
      updated++;
    }
    searched++;
    await sleep(DELAY_MS);
  }

  log(`Uploading updated Excel... (${updated} updated)`);
  try {
    const updatedBuffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([updatedBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    await storage.deleteFile(STORAGE_BUCKET_ID, EXCEL_FILE_ID);
    await storage.createFile(STORAGE_BUCKET_ID, EXCEL_FILE_ID, blob, ['read(\"any\")']);
    log('Excel uploaded successfully');
  } catch (e) {
    error(`Failed to upload Excel: ${e.message}`);
    return res.json({ success: false, error: e.message }, 500);
  }

  return res.json({ success: true, searched, updated, skipped, date: today });
};
