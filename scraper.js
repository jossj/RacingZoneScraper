'use strict';

const puppeteer = require('puppeteer');
const ExcelJS   = require('exceljs');
const readline  = require('readline');
const path      = require('path');
const fs        = require('fs');
const os        = require('os');

const cheerio = require('cheerio');

const DEFAULT_OUTPUT   = 'C:\\tab\\scrape';
const DEFAULT_DELAY_MS = 1000;

// ---------------------------------------------------------------------------
// CLI args
// ---------------------------------------------------------------------------

const args = parseArgs(process.argv.slice(2));

function parseArgs(argv) {
  const opts = { output: DEFAULT_OUTPUT, delayMs: DEFAULT_DELAY_MS };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--output') opts.output  = argv[++i];
    if (argv[i] === '--delay')  opts.delayMs = Number(argv[++i]) * 1000;
  }
  return opts;
}

// ---------------------------------------------------------------------------
// Logging
// ---------------------------------------------------------------------------

function ts() { return new Date().toISOString().replace('T', ' ').slice(0, 19); }
const info  = (...a) => console.log(`${ts()} [INFO ]`, ...a);
const warn  = (...a) => console.log(`${ts()} [WARN ]`, ...a);
const error = (...a) => console.log(`${ts()} [ERROR]`, ...a);

// ---------------------------------------------------------------------------
// Prompt
// ---------------------------------------------------------------------------

function prompt(question) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise(resolve => rl.question(question, ans => { rl.close(); resolve(ans.trim()); }));
}

async function promptForUrl() {
  console.log('\n' + '='.repeat(60));
  console.log('  Ladbrokes Form Guide Scraper');
  console.log('='.repeat(60));
  console.log('Paste the Ladbrokes form guide URL and press Enter.');
  console.log('Example: https://ladbrokesform.com.au/form/ba59f0cc-ece5-47fa-9b61-31f8b37ea042');
  console.log();
  let url = '';
  while (!url.startsWith('http')) {
    url = await prompt('Form guide URL: ');
    if (!url.startsWith('http')) console.log('  Please enter a valid URL starting with http.');
  }
  return url;
}

// ---------------------------------------------------------------------------
// Output path
// ---------------------------------------------------------------------------

function resolveOutputPath(raw) {
  if (process.platform === 'win32') return raw;
  const lower = raw.toLowerCase();
  if (lower.startsWith('c:\\') || lower.startsWith('c:/')) {
    return path.join(os.homedir(), raw.slice(3).replace(/\\/g, '/'));
  }
  return raw.replace(/\\/g, '/');
}

async function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

// ---------------------------------------------------------------------------
// Page preparation helpers
// ---------------------------------------------------------------------------

async function scrollPage(page) {
  await page.evaluate(async () => {
    await new Promise(resolve => {
      let scrolled = 0;
      const timer = setInterval(() => {
        window.scrollBy(0, 600);
        scrolled += 600;
        if (scrolled >= document.body.scrollHeight) {
          clearInterval(timer);
          window.scrollTo(0, 0);
          resolve();
        }
      }, 100);
    });
  });
  await sleep(800);
}

async function dismissOverlays(page) {
  const sels = [
    '[class*="cookie"] button[class*="accept"]',
    '[class*="cookie"] button[class*="agree"]',
    '[class*="cookie"] button[class*="close"]',
    '[id*="cookie"] button',
    '[class*="consent"] button',
    '[class*="modal"] button[class*="close"]',
    '[aria-label*="close" i]',
    'button[class*="dismiss"]',
  ];
  for (const sel of sels) {
    try {
      const btn = await page.$(sel);
      if (btn) { await btn.click(); await sleep(400); }
    } catch { /* ignore */ }
  }
}

async function expandAllRunners(page) {
  const clicked = await page.evaluate(() => {
    let count = 0;
    const candidates = [
      ...document.querySelectorAll('[aria-expanded="false"]'),
      ...document.querySelectorAll('[class*="expand"]'),
      ...document.querySelectorAll('[class*="toggle"]'),
      ...document.querySelectorAll('[class*="show-more"]'),
      ...document.querySelectorAll('[class*="show-form"]'),
    ];
    const seen = new Set();
    for (const el of candidates) {
      if (seen.has(el)) continue;
      seen.add(el);
      try { el.click(); count++; } catch { /* ignore */ }
    }
    return count;
  });
  if (clicked > 0) { await sleep(1500); }
}

// ---------------------------------------------------------------------------
// HTML snapshot + cheerio extraction
// ---------------------------------------------------------------------------

async function saveHtmlSnapshot(page, outputPath) {
  const html = await page.content();
  fs.mkdirSync(outputPath, { recursive: true });
  const htmlPath = path.join(outputPath, 'page_snapshot.html');
  fs.writeFileSync(htmlPath, html, 'utf8');
  info(`Page HTML saved: ${htmlPath} (${(html.length / 1024).toFixed(0)} KB)`);
  return html;
}

function extractRaceInfoFromHtml($) {
  const text = $('body').text();
  const info = {};

  for (const sel of ['h1','h2','[class*="race-name" i]','[class*="race-title" i]']) {
    const el = $(sel).first();
    if (el.length) { const t = el.text().trim(); if (t.length > 2 && t.length < 80) { info.raceName = t; break; } }
  }

  const raceNumM = text.match(/\bRace\s*(\d+)\b/i);
  if (raceNumM) info.raceNum = raceNumM[1];

  const distM = text.match(/\b(\d{3,4}m)\b/i);
  if (distM) info.distance = distM[1];

  const condM = text.match(/\b(Firm\s*\d*|Good\s*\d*|Soft\s*\d*|Heavy\s*\d*|Synthetic|Wet\s*\d*)\b/i);
  if (condM) info.trackCondition = condM[1].trim();

  const clsM = text.match(/\b(G[123]|Group\s*[123]|Listed|BM\s*\d+|Benchmark\s*\d+|MDN|Maiden|Handicap|Open|WFA|CL\d+)\b/i);
  if (clsM) info.raceClass = clsM[1].trim();

  for (const pat of [
    /\b(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{4})\b/i,
    /\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})\b/,
    /\b(\d{4}-\d{2}-\d{2})\b/,
  ]) {
    const m = text.match(pat);
    if (m) { info.date = m[1]; break; }
  }

  return info;
}

// Find runner containers using labeled fields as the racing signal.
// Horse names on this site are title-case (e.g. "Itchintogo"), not ALL CAPS.
function findRunnerContainers($) {
  const results = [];
  const seen = new WeakSet();

  $('div, section, article, li, ul').each(function() {
    const $el = $(this);
    const text = $el.text();

    if (text.trim().length < 30) return;   // too short
    if (text.length > 8000) return;        // wrapper element

    // Racing signals: labeled fields that appear in every runner card
    const hasRacingSignal =
      /Trainer\s*:/i.test(text) ||
      /Jockey\s*:/i.test(text)  ||
      /Prize Money\s*:/i.test(text);

    if (!hasRacingSignal) return;

    // Skip if a child also passes — keep the innermost match
    let childMatches = false;
    $el.children().each(function() {
      const ct = $(this).text();
      if (/Trainer\s*:/i.test(ct) || /Jockey\s*:/i.test(ct)) {
        childMatches = true;
        return false;
      }
    });
    if (childMatches) return;

    if (!seen.has(this)) { seen.add(this); results.push(this); }
  });

  return results.length >= 2 && results.length <= 30 ? results : [];
}

// Parse a single runner element into structured data.
// Expected list format (each item is a <li>):
//   "HorseName (barrier) formString"
//   "Trainer: Name(Venue)"
//   "Jockey: Name (Last 50: W-P-L)"
//   "Weight: 57kg"
//   "Prize Money: $237,300"
//   "Win / Place: 40% / 60%"
function extractRunnerFromElement($, el) {
  const $el = $(el);
  const text = $el.text();

  if (!text || text.trim().length < 10) return null;

  // ── Number from id attribute ──────────────────────────────────────────────
  const idNum = ($el.attr('id') || '').replace(/\D/g, '');

  // ── Collect all <li> texts for labeled-field extraction ───────────────────
  const liTexts = $el.find('li').map((_, li) => $(li).text().trim()).toArray();

  // ── First <li>: "HorseName (barrier) formString" ─────────────────────────
  const firstLi = liTexts[0] || '';
  // Match: one or more title-case words, then (number), then digit/x sequence
  const nbfMatch = firstLi.match(/^(.+?)\s+\((\d{1,2})\)\s+([0-9Xx]+)/);
  let name    = nbfMatch ? nbfMatch[1].trim() : firstLi.replace(/\s*\(\d+\).*/, '').trim();
  const barrier = nbfMatch ? nbfMatch[2] : '';
  const form    = nbfMatch ? nbfMatch[3] : '';

  if (!name || name.length < 2) return null;

  // ── Labeled fields ────────────────────────────────────────────────────────
  // Strip parenthesised extras like "(Eagle Farm)" or "(Last 50: 9-5-3)"
  const labelVal = (label) => {
    const li = liTexts.find(t => new RegExp(`^${label}\\s*:`, 'i').test(t));
    if (!li) return '';
    return li.replace(new RegExp(`^${label}\\s*:\\s*`, 'i'), '')
             .replace(/\s*\([^)]*\)\s*$/, '')  // strip trailing (...)
             .trim();
  };

  const trainer    = labelVal('Trainer');
  const weight     = labelVal('Weight');
  const prizeMoney = labelVal('Prize Money');

  // Jockey — strip trailing "(Last 50: ...)" stats
  const jockeyRaw = labelVal('Jockey');
  const jockey    = jockeyRaw.replace(/\s*\(Last\s+\d+[^)]*\)/i, '').trim();

  // Win / Place: "40% / 60%"
  const winPlaceLi = liTexts.find(t => /Win\s*\/\s*Place\s*:/i.test(t)) || '';
  const wpMatch    = winPlaceLi.match(/([\d.]+%)\s*\/\s*([\d.]+%)/);
  const winPct     = wpMatch ? wpMatch[1] : '';
  const placePct   = wpMatch ? wpMatch[2] : '';

  // ── Career stats ──────────────────────────────────────────────────────────
  // Pattern like "25: 5-4-3" or "25 5 4 3" somewhere in the text
  let careerStarts = '', careerWins = '', careerSeconds = '', careerThirds = '';
  const cm = text.match(/(\d{1,3})\s*[:\-]\s*(\d{1,3})\s*[:\-]\s*(\d{1,3})\s*[:\-]\s*(\d{1,3})/);
  if (cm) { careerStarts = cm[1]; careerWins = cm[2]; careerSeconds = cm[3]; careerThirds = cm[4]; }

  // ── Age/sex/colour ────────────────────────────────────────────────────────
  const asm = text.match(/(\d+yo\s+(?:Bay|Brown|Chestnut|Grey|Black|Roan|Palomino|White)\s+(?:Gelding|Mare|Colt|Filly|Stallion|Horse))/i);
  const ageSexColour = asm ? asm[1] : '';

  // ── Sire / Dam ────────────────────────────────────────────────────────────
  const sire = labelVal('Sire');
  const dam  = labelVal('Dam');

  // ── Odds — dollar amounts ─────────────────────────────────────────────────
  const allOdds = [...text.matchAll(/\$([\d.]+)/g)].map(m => m[1]).filter(v => v.includes('.'));
  const winOdds   = allOdds[0] || '';
  const placeOdds = allOdds[1] || '';

  // ── Race history ──────────────────────────────────────────────────────────
  const raceHistory = extractRaceHistoryFromElement($, $el);

  // ── Scratched ─────────────────────────────────────────────────────────────
  const scratched = /\bscratched\b|\bSCR\b/i.test(text);

  return {
    number: idNum,
    name:   name.replace(/\s+/g, ' ').trim(),
    barrier, jockey, trainer, weight, form,
    winOdds, placeOdds, ageSexColour, sire, dam, scratched,
    careerStarts, careerWins, careerSeconds, careerThirds,
    prizeMoney, winPct, placePct,
    condStats: {}, raceHistory,
  };
}

function parseResult(text) {
  const out = { first: '', second: '', third: '' };
  if (!text) return out;
  const parts = text.split(/\s*\/\s*/);
  for (const part of parts) {
    const m = part.match(/^(\d+)\.\s*(.+)$/);
    if (!m) continue;
    if (m[1] === '1') out.first  = m[2].trim();
    if (m[1] === '2') out.second = m[2].trim();
    if (m[1] === '3') out.third  = m[2].trim();
  }
  return out;
}

function buildColMap(headers) {
  const map = {};
  headers.forEach((c, i) => {
    const n = c.toLowerCase().replace(/[^a-z0-9]/g, '');
    if      (n === 'place')                        map.position   = i;
    else if (n === 'date')                         map.date       = i;
    else if (n === 'track')                        map.venue      = i;
    else if (n === 'distance')                     map.distance   = i;
    else if (n === 'class')                        map.raceClass  = i;
    else if (n === 'barrier')                      map.barrier    = i;
    else if (n === 'jockey')                       map.jockey     = i;
    else if (n === 'weight')                       map.weight     = i;
    else if (n === 'cond')                         map.condition  = i;
    else if (n === 'racetime')                     map.time       = i;
    else if (n === 'avgspeed')                     map.avgSpeed   = i;
    else if (n === 'last600m')                     map.last600m   = i;
    else if (n === 'margin')                       map.margin     = i;
    else if (n === '800400')                       map.sectionals = i;
    else if (n === 'sp')                           map.sp         = i;
    else if (n === 'days')                         map.days       = i;
    else if (n === 'prize' && map.prize === undefined) map.prize  = i;
    else if (n === 'prizewon')                     map.prizeWon   = i;
    else if (n === 'result' || n === 'placegetters') map.result   = i;
  });
  return map;
}

function cellsToHistRow(cells, colMap) {
  const get = f => (colMap[f] !== undefined ? cells[colMap[f]] || '' : '');
  const placement = parseResult(get('result'));
  return {
    position:   get('position'),
    date:       get('date'),
    venue:      get('venue'),
    distance:   get('distance'),
    raceClass:  get('raceClass'),
    barrier:    get('barrier'),
    jockey:     get('jockey'),
    weight:     get('weight'),
    condition:  get('condition'),
    time:       get('time'),
    avgSpeed:   get('avgSpeed'),
    last600m:   get('last600m'),
    margin:     get('margin'),
    sectionals: get('sectionals'),
    sp:         get('sp'),
    days:       get('days'),
    prize:      get('prize'),
    prizeWon:   get('prizeWon'),
    first:      placement.first,
    second:     placement.second,
    third:      placement.third,
    raw:        cells.join(' | '),
  };
}

function isHistHeader(cells) {
  return cells.some(c => /^place$/i.test(c) || /^track$/i.test(c) || /^date$/i.test(c));
}

function isDaySpell(cells) {
  return /\d+\s+Day\s+Spell/i.test(cells.join(' '));
}

function extractRaceHistoryFromElement($, $el) {
  const history = [];

  // ── Table-based extraction with header-driven column mapping ──────────────
  $el.find('table').each(function() {
    let colMap = null;
    const tableRows = [];
    $(this).find('tr').each(function() {
      const cells = $(this).find('td, th').map((_, c) => $(c).text().trim()).toArray();
      if (!cells.length || cells.every(c => !c)) return;
      if (isDaySpell(cells)) return;
      if (isHistHeader(cells)) { colMap = buildColMap(cells); return; }
      if (!colMap) return;
      if (isHistHeader(cells)) return;  // repeated header row
      const row = cellsToHistRow(cells, colMap);
      if (row.date || row.position) tableRows.push(row);
    });
    history.push(...tableRows);
  });
  if (history.length) return history;

  // ── Div-based extraction: find a header div then map positional children ──
  let colMap = null;
  $el.find('div, ul, ol').each(function() {
    const $row = $(this);
    if ($row.find('div').length > 12) return;  // skip wrapper elements
    const children = $row.children().toArray();
    if (children.length < 4) return;
    const cells = children.map(c => $(c).text().trim());
    if (cells.every(c => !c)) return;
    if (isDaySpell(cells)) return;
    if (isHistHeader(cells)) { colMap = buildColMap(cells); return; }
    if (!colMap) return;
    const row = cellsToHistRow(cells, colMap);
    if (row.date || row.position) history.push(row);
  });

  return history;
}

function extractRunnersFromHtml(html) {
  const $ = cheerio.load(html);

  // ── Strategy 1: ID-based selectors ────────────────────────────────────────
  const idPatterns = [
    '[id^="runner-"]', '[id^="Runner-"]',
    '[id^="competitor-"]', '[id^="horse-"]',
    '[id^="entry-"]', '[id^="field-"]',
  ];
  for (const pat of idPatterns) {
    const els = $(pat).toArray();
    if (els.length >= 2) {
      info(`cheerio: found ${els.length} elements via ${pat}`);
      const runners = els.map(el => extractRunnerFromElement($, el)).filter(r => r && r.name.length >= 2);
      if (runners.length >= 2) return { runners, raceInfo: extractRaceInfoFromHtml($) };
    }
  }

  // ── Strategy 2: Structural analysis ───────────────────────────────────────
  const els = findRunnerContainers($);
  if (els.length >= 2) {
    info(`cheerio: found ${els.length} runner containers via structural analysis`);
    const runners = els.map(el => extractRunnerFromElement($, el)).filter(r => r && r.name.length >= 2);
    if (runners.length >= 2) return { runners, raceInfo: extractRaceInfoFromHtml($) };
  }

  warn('cheerio: no runner containers found');
  return { runners: [], raceInfo: extractRaceInfoFromHtml($) };
}

// ---------------------------------------------------------------------------
// Excel colours / style helpers
// ---------------------------------------------------------------------------

const HEADER_BG  = 'FF1B5E20';
const SECTION_BG = 'FF388E3C';
const ALT_BG     = 'FFE8F5E9';
const WHITE      = 'FFFFFFFF';
const GREY_TEXT  = 'FF999999';

const thinBorder = {
  top: { style: 'thin' }, left: { style: 'thin' },
  bottom: { style: 'thin' }, right: { style: 'thin' },
};

function styleHeader(row, numCols) {
  for (let c = 1; c <= numCols; c++) {
    const cell = row.getCell(c);
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_BG } };
    cell.font      = { color: { argb: WHITE }, bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border    = thinBorder;
  }
  row.height = 22;
}

function styleSection(row, numCols) {
  for (let c = 1; c <= numCols; c++) {
    const cell = row.getCell(c);
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: SECTION_BG } };
    cell.font      = { color: { argb: WHITE }, bold: true };
    cell.alignment = { horizontal: 'left', vertical: 'middle' };
    cell.border    = thinBorder;
  }
}

function styleData(row, numCols, alternate = false) {
  const bg = alternate ? ALT_BG : 'FFFFFFFF';
  for (let c = 1; c <= numCols; c++) {
    const cell = row.getCell(c);
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    cell.border    = thinBorder;
  }
}

function autoWidth(ws, min = 8, max = 40) {
  ws.columns.forEach(col => {
    let maxLen = 0;
    col.eachCell({ includeEmpty: false }, cell => {
      const len = cell.value ? String(cell.value).length : 0;
      if (len > maxLen) maxLen = len;
    });
    col.width = Math.min(Math.max(maxLen + 2, min), max);
  });
}

// ---------------------------------------------------------------------------
// Excel export
// ---------------------------------------------------------------------------

async function saveToExcel(raceInfo, runners, outputPath) {
  fs.mkdirSync(outputPath, { recursive: true });

  const safeVenue = (raceInfo.raceName || raceInfo.venue || 'Ladbrokes')
    .replace(/[^a-zA-Z0-9 _-]/g, '_').replace(/\s+/g, '_').slice(0, 30);
  const dateStr = (raceInfo.date || new Date().toISOString().slice(0, 10))
    .replace(/[\s\/\-]/g, '').slice(0, 8);
  const raceTag = raceInfo.raceNum ? `_R${raceInfo.raceNum}` : '';
  const filename = `Ladbrokes_${safeVenue}${raceTag}_${dateStr}.xlsx`;
  const filepath = path.join(outputPath, filename);

  info('Saving Excel:', filepath);
  const wb = new ExcelJS.Workbook();

  // Sheet 1: Race Info
  {
    const ws = wb.addWorksheet('Race Info');
    ws.columns = [{ width: 22 }, { width: 55 }];
    styleHeader(ws.addRow(['Field', 'Value']), 2);
    [
      ['Race Name / Venue',   raceInfo.raceName       || raceInfo.venue || ''],
      ['Date',                raceInfo.date           || ''],
      ['Race Number',         raceInfo.raceNum        || ''],
      ['Distance',            raceInfo.distance       || ''],
      ['Class',               raceInfo.raceClass      || ''],
      ['Track Condition',     raceInfo.trackCondition || ''],
      ['Details',             raceInfo.details        || ''],
      ['URL',                 raceInfo.url            || ''],
      ['Scraped At',          new Date().toISOString().replace('T', ' ').slice(0, 19)],
    ].forEach(([k, v], i) => styleData(ws.addRow([k, v]), 2, i % 2 === 1));
  }

  // Sheet 2: Runners
  {
    const ws = wb.addWorksheet('Runners');
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    const cols = ['No.','Horse','Barrier','Jockey','Trainer','Weight','Form','Win Odds','Place Odds','Age/Sex/Colour','Sire','Dam','Scratched'];
    styleHeader(ws.addRow(cols), cols.length);
    runners.forEach((r, i) => {
      const row = ws.addRow([
        r.number, r.name, r.barrier, r.jockey, r.trainer, r.weight, r.form,
        r.winOdds, r.placeOdds, r.ageSexColour, r.sire, r.dam,
        r.scratched ? 'Yes' : 'No',
      ]);
      styleData(row, cols.length, i % 2 === 1);
      if (r.scratched) {
        for (let c = 1; c <= cols.length; c++)
          row.getCell(c).font = { color: { argb: GREY_TEXT }, italic: true };
      }
    });
    autoWidth(ws);
  }

  // Sheet 3: Career Stats
  {
    const ws = wb.addWorksheet('Career Stats');
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    const cols = ['No.','Horse','Starts','Wins','2nds','3rds','Prize Money','Win %','Place %'];
    styleHeader(ws.addRow(cols), cols.length);
    runners.forEach((r, i) => {
      styleData(ws.addRow([
        r.number, r.name,
        r.careerStarts, r.careerWins, r.careerSeconds, r.careerThirds,
        r.prizeMoney, r.winPct, r.placePct,
      ]), cols.length, i % 2 === 1);
    });
    autoWidth(ws);
  }

  // Sheet 4: Condition Stats
  {
    const ws = wb.addWorksheet('Condition Stats');
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    const allKeys = [...new Set(runners.flatMap(r => Object.keys(r.condStats || {})))];
    const cols = ['No.', 'Horse', ...allKeys];
    styleHeader(ws.addRow(cols), cols.length);
    runners.forEach((r, i) => {
      styleData(ws.addRow([
        r.number, r.name,
        ...allKeys.map(k => (r.condStats || {})[k] || ''),
      ]), cols.length, i % 2 === 1);
    });
    autoWidth(ws);
  }

  // Sheet 5: Race History
  {
    const ws = wb.addWorksheet('Race History');
    ws.views = [{ state: 'frozen', ySplit: 1, xSplit: 2 }];
    const cols = [
      'No.','Horse','Position','Date','Track','Distance','Class',
      'Barrier','Jockey','Weight','Condition','Time','Avg Speed',
      'Last 600m','Margin','800/400','SP','Days','Prize','Prize Won',
      '1st','2nd','3rd','Raw',
    ];
    styleHeader(ws.addRow(cols), cols.length);
    for (const r of runners) {
      if (!r.raceHistory || !r.raceHistory.length) continue;
      styleSection(ws.addRow([r.number, r.name, ...Array(cols.length - 2).fill('')]), cols.length);
      let alt = false;
      for (const e of r.raceHistory) {
        styleData(ws.addRow([
          r.number, r.name,
          e.position || '', e.date || '', e.venue || '', e.distance || '',
          e.raceClass || '', e.barrier || '', e.jockey || '', e.weight || '',
          e.condition || '', e.time || '', e.avgSpeed || '', e.last600m || '',
          e.margin || '', e.sectionals || '', e.sp || '', e.days || '',
          e.prize || '', e.prizeWon || '',
          e.first || '', e.second || '', e.third || '',
          e.raw || '',
        ]), cols.length, alt);
        alt = !alt;
      }
    }
    autoWidth(ws);
  }

  await wb.xlsx.writeFile(filepath);
  info('Excel saved:', filepath);
  return filepath;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  const url        = await promptForUrl();
  const outputPath = resolveOutputPath(args.output);
  info('Output directory:', outputPath);

  info('Launching Chrome...');
  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: null,
    args: ['--start-maximized'],
  });

  try {
    const [page] = await browser.pages();

    info('Navigating to:', url);
    await page.goto(url, { waitUntil: 'networkidle2', timeout: 60000 });
    await sleep(2000);

    await dismissOverlays(page);
    await scrollPage(page);
    await expandAllRunners(page);
    await sleep(1500);

    // ── Save rendered HTML and extract via cheerio ────────────────────────
    info('Saving rendered HTML snapshot...');
    const snapshotHtml = await saveHtmlSnapshot(page, outputPath);

    const cheerioResult = extractRunnersFromHtml(snapshotHtml);
    let runners  = cheerioResult.runners;
    let raceInfo = { url, ...cheerioResult.raceInfo };

    // ── Bail out if nothing found ─────────────────────────────────────────
    if (!runners.length) {
      error('No runners found. Open page_snapshot.html in the output folder to inspect the page structure.');
      process.exit(1);
    }

    const active    = runners.filter(r => !r.scratched);
    const scratched = runners.filter(r => r.scratched);
    info(`Runners: ${runners.length} total (${active.length} active, ${scratched.length} scratched)`);
    runners.forEach(r =>
      info(`  #${r.number || '?'} ${r.name} | J: ${r.jockey || '–'} | T: ${r.trainer || '–'} | Win: ${r.winOdds || '–'}`)
    );

    info(`Race: ${raceInfo.raceName || raceInfo.venue || 'Unknown'} | Date: ${raceInfo.date || 'N/A'} | Dist: ${raceInfo.distance || 'N/A'} | Cond: ${raceInfo.trackCondition || 'N/A'}`);

    const outputFile = await saveToExcel(raceInfo, runners, outputPath);
    console.log(`\nDone! Excel file saved to: ${outputFile}`);

  } finally {
    await browser.close();
  }
}

main().catch(err => {
  error('Fatal error:', err.message);
  process.exit(1);
});
