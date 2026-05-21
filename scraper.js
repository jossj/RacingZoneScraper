'use strict';

const puppeteer = require('puppeteer');
const ExcelJS   = require('exceljs');
const readline  = require('readline');
const path      = require('path');
const fs        = require('fs');
const os        = require('os');

const RACINGZONE_HORSES_URL = 'https://www.racingzone.com.au/statistics/horses/';
const DEFAULT_OUTPUT        = 'C:\\tab\\scrape';
const DEFAULT_DELAY_MS      = 2000;

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
  console.log('  Horse Racing Scraper');
  console.log('='.repeat(60));
  console.log('Paste the full URL of the TAB race page and press Enter.');
  console.log('Example: https://www.tab.com.au/racing/2026-05-18/RANDWICK/NSW/R/1');
  console.log();

  let url = '';
  while (!url.startsWith('http')) {
    url = await prompt('TAB race URL: ');
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

// ---------------------------------------------------------------------------
// TAB helpers
// ---------------------------------------------------------------------------

async function waitForRunners(page, timeout = 30000) {
  await page.waitForSelector('.runner-name', { timeout });
}

async function clickShowAllForm(page) {
  const sel = "[class*='show-all-form'] button, .show-all-form-wrapper button";
  try {
    await page.waitForSelector(sel, { timeout: 10000 });
    const btn = await page.$(sel);
    if (btn) {
      const text = await page.evaluate(el => el.textContent, btn);
      if (!text.includes('Hide')) {
        await btn.click();
        await sleep(4000);
        info('"Show All Form" expanded.');
        return;
      }
      info('Form already expanded — skipping click.');
    }
  } catch {
    warn('"Show All Form" button not found — trying individual expand buttons.');
    await expandIndividualForms(page);
  }
}

async function expandIndividualForms(page) {
  try {
    const btns = await page.$$(
      "[class*='form-toggle'] button, [class*='expand-form'] button"
    );
    for (const btn of btns) {
      try {
        const text = await page.evaluate(el => el.textContent, btn);
        if (!text.includes('Hide')) {
          await page.evaluate(el => el.click(), btn);
          await sleep(300);
        }
      } catch { /* ignore */ }
    }
    await sleep(2000);
    info(`Expanded ${btns.length} individual form panels.`);
  } catch { /* ignore */ }
}

async function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

// ---------------------------------------------------------------------------
// Race info
// ---------------------------------------------------------------------------

function parseUrlMeta(url) {
  const meta = { url };
  // e.g. /racing/2026-05-18/RANDWICK/NSW/R/3
  const m = url.match(/racing\/(\d{4}-\d{2}-\d{2})\/([A-Z0-9_-]+)\/[A-Z]+\/[RGH]\/(\d+)/i);
  if (m) { meta.date = m[1]; meta.venue = m[2].replace(/-/g, ' '); meta.raceNum = m[3]; }
  return meta;
}

async function getRaceInfo(page) {
  const url  = page.url();
  const meta = parseUrlMeta(url);

  Object.assign(meta, await page.evaluate(() => {
    const result = { raceName: '', bannerDetails: '', trackCondition: '' };

    // Race name
    for (const h of document.querySelectorAll('h1,h2,h3,h4')) {
      const t = h.textContent.trim();
      if (t.length > 5 && !/TAB|Racing/i.test(t)) {
        result.raceName = t.replace(/\s*-\s*Betting Odds$/i, '').trim();
        break;
      }
    }

    // Banner
    result.bannerDetails = [...document.querySelectorAll(
      "banner li, [class*='race-info'] li, [class*='race-detail'] li"
    )].map(e => e.textContent.trim()).filter(Boolean).slice(0, 10).join(' | ');

    // Track condition
    result.trackCondition = [...document.querySelectorAll(
      "[class*='track-condition'], [class*='condition']"
    )].map(e => e.textContent.trim()).filter(Boolean).join(' ');

    return result;
  }));

  return meta;
}

// ---------------------------------------------------------------------------
// Runners
// ---------------------------------------------------------------------------

async function scrapeRunners(page) {
  return page.evaluate(() => {
    const rows = [...document.querySelectorAll('.row')]
      .filter(r => r.querySelector('.runner-name'));

    return rows.map(row => {
      // Number / cloth
      const numEl = row.querySelector('.number-cell, [class*="number"], [class*="silk"]');
      const number = numEl ? numEl.textContent.trim() : '';

      // Horse name — strip nested span text
      let name = '';
      const nameEl = row.querySelector('.runner-name');
      if (nameEl) {
        const nested = [...nameEl.querySelectorAll('span, .barrier, .box')]
          .map(n => n.textContent).join('');
        name = nameEl.textContent.replace(nested, '').trim();
      }

      // Jockey + Trainer
      const fullNames = [...row.querySelectorAll('.runner-metadata-list .full-name')];
      const jockey  = fullNames[0] ? fullNames[0].textContent.trim() : '';
      const trainer = fullNames[1] ? fullNames[1].textContent.trim() : '';

      // Form / Weight / Rating
      const dts = [...row.querySelectorAll('.runner-metadata-list.optional dt')];
      const dds = [...row.querySelectorAll('.runner-metadata-list.optional dd')];
      const meta = {};
      dts.forEach((k, i) => { if (dds[i]) meta[k.textContent.trim()] = dds[i].textContent.trim(); });

      // Odds
      const prices = [...row.querySelectorAll('.price-cell')].map(p => p.textContent.trim());

      const winOdds   = prices[0] || '';
      const placeOdds = prices[1] || '';

      return {
        number, name, jockey, trainer,
        form:       meta['F'] || '',
        weight:     meta['W'] || '',
        rating:     meta['R'] || '',
        winOdds,
        placeOdds,
        toteWin:    prices[2] || '',
        totePlace:  prices[3] || '',
        scratched:  (winOdds + ' ' + placeOdds).includes('SCR'),
      };
    });
  });
}

// ---------------------------------------------------------------------------
// Form data (expanded panels)
// ---------------------------------------------------------------------------

async function scrapeFormData(page) {
  return page.evaluate(() => {
    const formRows = [...document.querySelectorAll('.row.form')];

    return formRows.map(formRow => {
      const wrapper = formRow.querySelector('.form-data-wrapper');
      if (!wrapper) return null;

      // Horse name from preceding sibling runner row
      let name = 'Unknown';
      let el = formRow.previousElementSibling;
      while (el && !el.querySelector('.runner-name')) el = el.previousElementSibling;
      if (el) name = el.querySelector('.runner-name').textContent.trim();

      // li text helper
      const liTexts = [...wrapper.querySelectorAll('li')]
        .map(li => li.textContent.trim()).filter(Boolean);
      const findVal = key => {
        const item = liTexts.find(t => t.startsWith(key));
        return item ? item.slice(key.length).trim() : '';
      };

      const winsItem   = liTexts.find(t => t.includes('Wins')   && t.includes('%')) || '';
      const placesItem = liTexts.find(t => t.includes('Places') && t.includes('%')) || '';
      const winsPct    = (winsItem.match(/(\d+)%/)   || [])[1];
      const placesPct  = (placesItem.match(/(\d+)%/) || [])[1];

      const profile = {
        career:     findVal('Career '),
        prizeMoney: findVal('Prize Money '),
        sire:       findVal('Sire '),
        dam:        findVal('Dam '),
        colour:     findVal('Colour '),
        sex:        findVal('Sex '),
        age:        findVal('Age '),
        trainer:    findVal('Trainer '),
        jockey:     findVal('Jockey '),
        owner:      findVal('Owner '),
        lastRun:    findVal('Last Run '),
        winsPct:    winsPct   ? winsPct   + '%' : '',
        placesPct:  placesPct ? placesPct + '%' : '',
      };

      const conditionStats = {
        track:    findVal('Track '),
        distance: findVal('Distance '),
        trkDist:  findVal('Trk & Dist '),
        firm:     findVal('Firm '),
        good:     findVal('Good '),
        soft:     findVal('Soft '),
        heavy:    findVal('Heavy '),
        barrier:  findVal('Barrier '),
        firstUp:  findVal('1st Up '),
        secondUp: findVal('2nd Up '),
        thirdUp:  findVal('3rd Up '),
      };

      // Race history
      const raceHistory = [];
      const bodyWrapper = wrapper.querySelector('.flexible-body-wrapper');
      if (bodyWrapper) {
        for (const bRow of bodyWrapper.querySelectorAll('.flexible-row')) {
          const spellEl = bRow.querySelector('.runner-spell .message');
          if (spellEl) {
            raceHistory.push({ type: 'spell', message: spellEl.textContent.trim() });
            continue;
          }
          const cells = [...bRow.querySelectorAll('.flexible-cell')].map(c => c.textContent.trim());
          if (cells.length >= 8 && (cells[0] || cells[2])) {
            raceHistory.push({
              type:      'race',
              placing:   cells[0]  || '',
              venue:     cells[1]  || '',
              date:      cells[2]  || '',
              class:     cells[3]  || '',
              distance:  cells[4]  || '',
              weight:    cells[5]  || '',
              barrier:   cells[6]  || '',
              odds:      cells[7]  || '',
              winner2nd: cells[8]  || '',
              margin:    cells[9]  || '',
              time:      cells[10] || '',
              inRun:     cells[11] || '',
            });
          }
        }
      }

      return { name, profile, conditionStats, raceHistory };
    }).filter(Boolean);
  });
}

// ---------------------------------------------------------------------------
// RacingZone
// ---------------------------------------------------------------------------

async function scrapeRacingZoneHorse(page, horseName) {
  const stats = {
    name: horseName, error: '',
    careerStarts: '', careerWins: '', careerSeconds: '', careerThirds: '',
    careerWinPct: '', careerPlacePct: '', careerPrizeMoney: '',
    l12mStarts: '', l12mWins: '', l12mSeconds: '', l12mThirds: '',
    statsByDistance: '', statsByCondition: '', statsByTrackType: '',
    statsByJockey: '', statsByTrainer: '',
  };

  info(`Searching RacingZone for: ${horseName}`);
  try {
    await page.goto(RACINGZONE_HORSES_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await sleep(2000);

    const inputSels = [
      "input[name='horse_name']", "input[name='name']",
      "input[placeholder*='horse' i]", "input[placeholder*='Find' i]",
      "input[type='search']", "#horse_name", "#name", "form input[type='text']",
    ];
    let input = null;
    for (const sel of inputSels) {
      input = await page.$(sel);
      if (input) break;
    }
    if (!input) { stats.error = 'Search input not found'; return stats; }

    await input.click({ clickCount: 3 });
    await input.type(horseName);
    await sleep(500);

    const btnSels = [
      "button[type='submit']", "input[type='submit']",
      "button[class*='search' i]", "form button",
    ];
    let clicked = false;
    for (const sel of btnSels) {
      const btn = await page.$(sel);
      if (btn) { await btn.click(); clicked = true; break; }
    }
    if (!clicked) await input.press('Enter');
    await sleep(3000);

    // Click best-matching result link
    const links = await page.$$('[class*="result"] a, table a, main a, #content a');
    if (links.length) {
      const nl = horseName.toLowerCase();
      let found = false;
      for (const link of links) {
        const txt = (await page.evaluate(el => el.textContent, link)).toLowerCase();
        if (txt.includes(nl) || nl.includes(txt)) {
          await link.click();
          found = true;
          await sleep(3000);
          break;
        }
      }
      if (!found) { await links[0].click(); await sleep(3000); }
    }

    await parseHorseStatsPage(page, stats);
  } catch (err) {
    error(`RacingZone error for ${horseName}:`, err.message);
    stats.error = err.message.slice(0, 200);
  }
  return stats;
}

async function parseHorseStatsPage(page, stats) {
  const pageText = await page.evaluate(() => document.body.innerText);
  if (!pageText.trim()) { stats.error = 'Empty page'; return; }

  await page.evaluate((s) => {
    const detailMap = {
      sire:    ['sire', 'father'],
      dam:     ['dam', 'mother'],
      colour:  ['colour', 'color'],
      sex:     ['sex', 'gender'],
      age:     ['age'],
      country: ['country', 'origin'],
      owner:   ['owner'],
      breeder: ['breeder'],
    };

    for (const tbl of document.querySelectorAll('table')) {
      for (const row of tbl.querySelectorAll('tr')) {
        const cells = [...row.querySelectorAll('td, th')];
        if (cells.length < 2) continue;
        const label = cells[0].textContent.toLowerCase().replace(/:$/, '');
        const value = cells[1].textContent.trim();
        for (const [attr, kws] of Object.entries(detailMap)) {
          if (kws.some(kw => label.includes(kw))) { s[attr] = value; break; }
        }
      }
    }

    const dts = [...document.querySelectorAll('dt')];
    const dds = [...document.querySelectorAll('dd')];
    for (let i = 0; i < Math.min(dts.length, dds.length); i++) {
      const label = dts[i].textContent.toLowerCase().replace(/:$/, '');
      const value = dds[i].textContent.trim();
      for (const [attr, kws] of Object.entries(detailMap)) {
        if (kws.some(kw => label.includes(kw))) { s[attr] = value; break; }
      }
    }

    return s;
  }, stats);

  // Stats tables
  const tableData = await page.evaluate(() => {
    return [...document.querySelectorAll('table')].map(tbl => ({
      headers: [...tbl.querySelectorAll('th')].map(h => h.textContent.toLowerCase()),
      rows: [...(tbl.querySelector('tbody') ? tbl.querySelectorAll('tbody tr') : [...tbl.querySelectorAll('tr')].slice(1))].map(row =>
        [...row.querySelectorAll('td')].map(td => td.textContent.trim())
      ),
    }));
  });

  for (const { headers, rows } of tableData) {
    for (const cells of rows) {
      if (!cells.length) continue;
      const label = cells[0].toLowerCase();
      if (label.includes('career') || label.includes('total') || label.includes('all')) {
        fillStats(stats, cells, headers, 'career');
      } else if (label.includes('12') || label.includes('l12') || label.includes('year')) {
        fillStats(stats, cells, headers, 'l12m');
      }
    }
  }
}

function fillStats(stats, cells, headers, prefix) {
  const mapping = {
    start: `${prefix}Starts`, win: `${prefix}Wins`,
    '2nd': `${prefix}Seconds`, second: `${prefix}Seconds`,
    '3rd': `${prefix}Thirds`,  third: `${prefix}Thirds`,
    'win%': `${prefix}WinPct`, 'place%': `${prefix}PlacePct`,
    prize: `${prefix}PrizeMoney`, earning: `${prefix}PrizeMoney`,
  };
  cells.slice(1).forEach((val, idx) => {
    const hdr = (headers[idx] || '').toLowerCase();
    for (const [key, attr] of Object.entries(mapping)) {
      if (hdr.includes(key) && attr in stats) { stats[attr] = val; break; }
    }
  });
}

// ---------------------------------------------------------------------------
// Excel colours / style helpers  (matches greyhoundracing/index.js palette)
// ---------------------------------------------------------------------------

const HEADER_BG  = 'FF1B5E20';
const SECTION_BG = 'FF388E3C';
const ALT_BG     = 'FFE8F5E9';
const SPELL_BG   = 'FFFFE699';
const SPELL_FG   = 'FF7F4F00';
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

function styleSpell(row, numCols) {
  for (let c = 1; c <= numCols; c++) {
    const cell = row.getCell(c);
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: SPELL_BG } };
    cell.font      = { color: { argb: SPELL_FG }, bold: true, italic: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
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

async function saveToExcel(raceInfo, runners, formData, horseStats, outputPath) {
  fs.mkdirSync(outputPath, { recursive: true });

  const safeVenue = (raceInfo.venue || 'Unknown').replace(/[^a-zA-Z0-9 _-]/g, '_');
  const dateStr   = (raceInfo.date  || new Date().toISOString().slice(0, 10)).replace(/-/g, '');
  const filename  = `TAB_${safeVenue}_R${raceInfo.raceNum || '0'}_${dateStr}.xlsx`.replace(/\s+/g, '_');
  const filepath  = path.join(outputPath, filename);

  info('Saving Excel:', filepath);
  const wb = new ExcelJS.Workbook();

  const stripNum = name => name.replace(/\s*\(\d+\)\s*$/, '').trim();

  // ── Sheet 1: Race Info ──────────────────────────────────────────────────
  const ws1 = wb.addWorksheet('Race Info');
  ws1.views = [{ state: 'frozen', ySplit: 1, xSplit: 1 }];
  ws1.columns = [{ width: 20 }, { width: 50 }];
  styleHeader(ws1.addRow(['Field', 'Value']), 2);
  [
    ['Race Name',       raceInfo.raceName       || ''],
    ['Date',            raceInfo.date           || ''],
    ['Venue',           raceInfo.venue          || ''],
    ['Race Number',     raceInfo.raceNum        || ''],
    ['Track Condition', raceInfo.trackCondition || ''],
    ['Banner Details',  raceInfo.bannerDetails  || ''],
    ['URL',             raceInfo.url            || ''],
    ['Scraped At',      new Date().toISOString().replace('T', ' ').slice(0, 19)],
  ].forEach(([k, v], i) => styleData(ws1.addRow([k, v]), 2, i % 2 === 1));

  // ── Sheet 2: Runners ───────────────────────────────────────────────────
  const ws2 = wb.addWorksheet('Runners (TAB)');
  ws2.views = [{ state: 'frozen', ySplit: 1 }];
  const runnerCols = [
    'No.', 'Horse', 'Jockey', 'Trainer', 'Form', 'Weight', 'Rating',
    'FO Win', 'FO Place', 'Tote Win', 'Tote Place', 'Scratched',
  ];
  styleHeader(ws2.addRow(runnerCols), runnerCols.length);
  runners.forEach((r, i) => {
    const row = ws2.addRow([
      r.number, r.name, r.jockey, r.trainer, r.form, r.weight, r.rating,
      r.winOdds, r.placeOdds, r.toteWin, r.totePlace,
      r.scratched ? 'Yes' : 'No',
    ]);
    styleData(row, runnerCols.length, i % 2 === 1);
    if (r.scratched) {
      for (let c = 1; c <= runnerCols.length; c++)
        row.getCell(c).font = { color: { argb: GREY_TEXT }, italic: true };
    }
  });
  autoWidth(ws2);

  // ── Sheet 3: Horse Profiles ────────────────────────────────────────────
  const ws3 = wb.addWorksheet('Horse Profiles');
  ws3.views = [{ state: 'frozen', ySplit: 1 }];
  const profCols = [
    'No.', 'Horse', 'Career', 'Prize Money', 'Sire', 'Dam',
    'Colour', 'Sex', 'Age', 'Owner', 'Trainer', 'Jockey',
    'Last Run', 'Wins %', 'Places %',
  ];
  styleHeader(ws3.addRow(profCols), profCols.length);
  const runnerMap = new Map(runners.map(r => [r.name.toUpperCase(), r]));
  formData.forEach((fd, i) => {
    const clean   = stripNum(fd.name);
    const runner  = runnerMap.get(clean.toUpperCase()) || {};
    const p       = fd.profile;
    const row = ws3.addRow([
      runner.number || '', clean,
      p.career, p.prizeMoney, p.sire, p.dam, p.colour, p.sex, p.age,
      p.owner, p.trainer || runner.trainer, p.jockey || runner.jockey,
      p.lastRun, p.winsPct, p.placesPct,
    ]);
    styleData(row, profCols.length, i % 2 === 1);
  });
  autoWidth(ws3);

  // ── Sheet 4: Condition Stats ───────────────────────────────────────────
  const ws4 = wb.addWorksheet('Condition Stats');
  ws4.views = [{ state: 'frozen', ySplit: 1 }];
  const condCols = [
    'No.', 'Horse', 'Track', 'Distance', 'Trk & Dist',
    'Firm', 'Good', 'Soft', 'Heavy', 'Barrier', '1st Up', '2nd Up', '3rd Up',
  ];
  styleHeader(ws4.addRow(condCols), condCols.length);
  formData.forEach((fd, i) => {
    const clean  = stripNum(fd.name);
    const runner = runnerMap.get(clean.toUpperCase()) || {};
    const cs     = fd.conditionStats;
    styleData(ws4.addRow([
      runner.number || '', clean,
      cs.track, cs.distance, cs.trkDist,
      cs.firm, cs.good, cs.soft, cs.heavy,
      cs.barrier, cs.firstUp, cs.secondUp, cs.thirdUp,
    ]), condCols.length, i % 2 === 1);
  });
  autoWidth(ws4);

  // ── Sheet 5: Race History ──────────────────────────────────────────────
  const ws5 = wb.addWorksheet('Race History');
  ws5.views = [{ state: 'frozen', ySplit: 1, xSplit: 2 }];
  const histCols = [
    'No.', 'Horse', 'Placing', 'Venue', 'Date', 'Class',
    'Dist', 'Weight', 'Barrier', 'Odds', 'Winner/2nd',
    'Margin', 'Time', 'In Run',
  ];
  styleHeader(ws5.addRow(histCols), histCols.length);

  let curRow = 2;
  for (const fd of formData) {
    const clean  = stripNum(fd.name);
    const runner = runnerMap.get(clean.toUpperCase()) || {};
    const num    = runner.number || '';

    const secRow = ws5.addRow([num, clean, ...Array(histCols.length - 2).fill('')]);
    styleSection(secRow, histCols.length);
    curRow++;

    let alt = false;
    for (const entry of fd.raceHistory) {
      if (entry.type === 'spell') {
        const r = ws5.addRow([entry.message]);
        ws5.mergeCells(curRow, 1, curRow, histCols.length);
        styleSpell(r, histCols.length);
        alt = false;
      } else {
        const r = ws5.addRow([
          num, clean, entry.placing, entry.venue, entry.date, entry.class,
          entry.distance, entry.weight, entry.barrier, entry.odds,
          entry.winner2nd, entry.margin, entry.time, entry.inRun,
        ]);
        styleData(r, histCols.length, alt);
        alt = !alt;
      }
      curRow++;
    }
  }
  autoWidth(ws5);

  // ── Sheet 6: RacingZone Stats ──────────────────────────────────────────
  const ws6 = wb.addWorksheet('RacingZone Stats');
  ws6.views = [{ state: 'frozen', ySplit: 1 }];
  const rzCols = [
    'Horse', 'Career Starts', 'Career Wins', 'Career 2nds', 'Career 3rds',
    'Career Win %', 'Career Place %', 'Career Prize $',
    'L12M Starts', 'L12M Wins', 'L12M 2nds', 'L12M 3rds',
    'By Distance', 'By Condition', 'By Track Type', 'By Jockey', 'By Trainer', 'Error',
  ];
  styleHeader(ws6.addRow(rzCols), rzCols.length);
  horseStats.forEach((s, i) => {
    styleData(ws6.addRow([
      s.name, s.careerStarts, s.careerWins, s.careerSeconds, s.careerThirds,
      s.careerWinPct, s.careerPlacePct, s.careerPrizeMoney,
      s.l12mStarts, s.l12mWins, s.l12mSeconds, s.l12mThirds,
      s.statsByDistance, s.statsByCondition, s.statsByTrackType,
      s.statsByJockey, s.statsByTrainer, s.error,
    ]), rzCols.length, i % 2 === 1);
  });
  autoWidth(ws6);

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
    defaultViewport: null,          // use real window size
    args: ['--start-maximized'],
  });

  try {
    // Reuse the tab Chrome already opened (same as greyhoundracing/index.js)
    const [page] = await browser.pages();

    info('Navigating to:', url);
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });

    info('Waiting for runners...');
    await waitForRunners(page);
    await sleep(2000);

    // Dismiss cookie banner
    try {
      const cookieBtn = await page.$("[class*='cookie'] button, [id*='cookie'] button");
      if (cookieBtn) { await cookieBtn.click(); await sleep(1000); info('Cookie banner dismissed.'); }
    } catch { /* ignore */ }

    info('Clicking "Show All Form"...');
    await clickShowAllForm(page);
    await sleep(3000);

    info('Scraping race info...');
    const raceInfo = await getRaceInfo(page);
    info(`Race: ${raceInfo.raceName || 'Unknown'} | ${raceInfo.date || ''} | ${raceInfo.venue || ''} | R${raceInfo.raceNum || ''}`);

    info('Scraping runners...');
    const runners   = await scrapeRunners(page);
    const active    = runners.filter(r => !r.scratched);
    const scratched = runners.filter(r => r.scratched);
    info(`Runners: ${runners.length} total (${active.length} active, ${scratched.length} scratched)`);

    if (!runners.length) {
      error('No runners found. Check the URL and try again. Page title:', await page.title());
      process.exit(1);
    }

    info('Scraping form data...');
    const formData = await scrapeFormData(page);
    info(`Form data scraped for ${formData.length} horses`);

    info('Starting RacingZone lookups...');
    const horseStats = [];
    for (let i = 0; i < runners.length; i++) {
      const runner = runners[i];
      if (runner.scratched) {
        horseStats.push({ name: runner.name, error: 'Scratched' });
        continue;
      }
      info(`[${i + 1}/${runners.length}] RacingZone: ${runner.name}`);
      horseStats.push(await scrapeRacingZoneHorse(page, runner.name));
      if (i < runners.length - 1) await sleep(args.delayMs);
    }

    const outputFile = await saveToExcel(raceInfo, runners, formData, horseStats, outputPath);
    console.log(`\nDone! Excel file saved to: ${outputFile}`);

  } finally {
    await browser.close();
  }
}

main().catch(err => {
  error('Fatal error:', err.message);
  process.exit(1);
});
