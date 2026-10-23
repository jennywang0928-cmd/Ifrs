// 從證交所 ESG InfoHub 的永續報告書 PDF 擷取「報告書聯絡資訊」那一頁的文字。
// PDF 很大、連線很慢，所以用 pdf.js + HTTP Range 只讀需要的頁面，依「最可能出現聯絡資訊」的順序逐頁找，找到就停。
// 用法: node extract_contacts.mjs <reports.json> <輸出.jsonl> [平行數]
// 需要: npm i pdfjs-dist@4.10.38 undici@6.21.1
import { getDocument, PDFDataRangeTransport } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { ProxyAgent, request } from 'undici';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';

const agent = process.env.HTTPS_PROXY ? new ProxyAgent(process.env.HTTPS_PROXY) : undefined;
const URLB = 'https://esggenplus.twse.com.tw/api/api/MopsSustainReport/data/FileStream?id=';
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36';
const ZERO = '00000000-0000-0000-0000-000000000000';
const PDFJS = path.dirname(createRequire(import.meta.url).resolve('pdfjs-dist/package.json'));
const KEY = /聯絡|聯繫|意見回饋|回饋意見|洽詢|窗口|聯系/;

const sleep = ms => new Promise(z => setTimeout(z, ms));

async function rng(url, s, e, stat) {
  for (let a = 0; a < 10; a++) {
    try {
      const r = await request(url, { dispatcher: agent, headers: { 'user-agent': UA, range: `bytes=${s}-${e}` }, headersTimeout: 90000, bodyTimeout: 120000 });
      const buf = Buffer.from(await r.body.arrayBuffer());
      if (!String(r.headers['content-type']).includes('pdf')) throw new Error('blocked');
      stat.bytes += buf.length;
      return { buf, cr: r.headers['content-range'] };
    } catch (err) {
      await sleep(Math.min(60000, 1500 * 2 ** a) + Math.random() * 1000);
    }
  }
  throw new Error(`range ${s}-${e} failed`);
}

const norm = s => (s || '').replace(/＠/g, '@').replace(/\s+/g, '').toLowerCase();

// DEEP=1：第二輪，給第一輪沒找到的報告書，改掃後面的頁面
const DEEP = process.env.DEEP === '1';

function order(n) {
  const want = DEEP
    ? [...Array.from({ length: 18 }, (_, i) => 13 + i), n - 3, n - 4, n - 5, n - 6]
    : [2, 3, 4, 5, 6, n, n - 1, 1, 7, 8, 9, 10, 11, 12, n - 2];
  return [...new Set(want.filter(p => p >= 1 && p <= n))];
}

async function processReport(r) {
  const e = r.twEditReportDownloadId;
  const fid = e && e !== ZERO ? e : r.twFirstReportDownloadId;
  const stat = { bytes: 0 };
  const res = { code: r.code, file_id: fid, pages: [], error: null };
  const t0 = Date.now();
  let doc;
  try {
    const url = URLB + fid;
    const len = +(await rng(url, 0, 0, stat)).cr.split('/')[1];
    class T extends PDFDataRangeTransport {
      requestDataRange(b, end) { rng(url, b, end - 1, stat).then(({ buf }) => this.onDataRange(b, new Uint8Array(buf))).catch(() => {}); }
    }
    doc = await getDocument({ range: new T(len, null), length: len, disableAutoFetch: true, disableStream: true, rangeChunkSize: 65536,
      cMapUrl: PDFJS + '/cmaps/', cMapPacked: true, standardFontDataUrl: PDFJS + '/standard_fonts/', isEvalSupported: false, verbosity: 0 }).promise;
    const emails = norm(r.contactInfo).match(/[a-z0-9._%+\-]+@[a-z0-9.\-]+/g) || [];
    res.n_pages = doc.numPages;
    let best = null;
    for (const p of order(doc.numPages)) {
      const page = await doc.getPage(p);
      const tc = await Promise.race([page.getTextContent(), sleep(300000).then(() => { throw new Error('page timeout'); })]);
      const text = tc.items.map(i => i.str + (i.hasEOL ? '\n' : '')).join('');
      const nt = norm(text);
      const hitEmail = emails.some(m => nt.includes(m));
      const score = (hitEmail ? 10 : 0) + (KEY.test(text) && /@|＠/.test(text) ? 5 : 0) + (KEY.test(text) ? 1 : 0);
      if (score && (!best || score > best.score)) best = { page: p, score, text: text.slice(0, 8000) };
      if (score >= 5) break;
      if (Date.now() - t0 > 15 * 60000) break;
    }
    if (best) res.pages.push(best);
  } catch (err) {
    res.error = String(err.message || err).slice(0, 300);
  } finally {
    try { await doc?.destroy(); } catch {}
  }
  res.bytes = stat.bytes;
  res.secs = Math.round((Date.now() - t0) / 1000);
  return res;
}

const [src, out, workersArg] = process.argv.slice(2);
const workers = +(workersArg || 4);
const all = Object.values(JSON.parse(fs.readFileSync(src, 'utf8'))).flatMap(v => v.data);
const done = new Set(fs.existsSync(out) ? fs.readFileSync(out, 'utf8').split('\n').filter(Boolean).map(l => JSON.parse(l).code) : []);
const todo = all.filter(r => !done.has(r.code));
console.log(`total ${all.length}, done ${done.size}, todo ${todo.length}`);
let k = 0;
await Promise.all(Array.from({ length: workers }, async () => {
  while (todo.length) {
    const r = todo.shift();
    const res = await processReport(r);
    fs.appendFileSync(out, JSON.stringify(res) + '\n');
    console.log(++k, res.code, res.error || `p${res.pages.map(p => p.page + ':' + p.score)} ${(res.bytes / 1024) | 0}KB ${res.secs}s`);
  }
}));
process.exit(0);
