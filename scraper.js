'use strict';

const readline = require('readline');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { chromium, firefox } = require('playwright');
const ExcelJS = require('exceljs');

const RACINGZONE_HORSES_URL = 'https://www.racingzone.com.au/statistics/horses/';
const DEFAULT_OUTPUT = 'C:\\tab\\scrape';
const DEFAULT_DELAY_MS = 2000;

// ---------------------------------------------------------------------------
// CLI args (optional flags only — URL is prompted interactively)
// ---------------------------------------------------------------------------

const args = parseArgs(process.argv.slice(2));

function parseArgs(argv) {
  const opts = {
    output: DEFAULT_OUTPUT,
    browser: 'chromium',
    headless: false,   // always show the browser window
    delayMs: DEFAULT_DELAY_MS,
  };
  for (let i = 0; i < argv.length; i++) {
    switch (argv[i]) {
      case '--output':  opts.output  = argv[++i]; break;
      case '--browser': opts.browser = argv[++i]; break;
      case '--delay':   opts.delayMs = Number(argv[++i]) * 1000; break;
    }
  }
  return opts;
}

// ---------------------------------------------------------------------------
// Logging
// ---------------------------------------------------------------------------

function log(level, ...parts) {
  const ts = new Date().toISOString().replace('T', ' ').slice(0, 19);
  console.log(`${ts} [${level}]`, ...parts);
}
const info  = (...a) => log('INFO ', ...a);
const warn  = (...a) => log('WARN ', ...a);
const error = (...a) => log('ERROR', ...a);
const debug = (...a) => { if (process.env.DEBUG) log('DEBUG', ...a); };

// ---------------------------------------------------------------------------
// Interactive URL prompt
// ---------------------------------------------------------------------------

function promptForUrl() {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    console.log('\n' + '='.repeat(60));
    console.log('  Horse Racing Scraper');
    console.log('='.repeat(60));
    console.log('Paste the full URL of the TAB race page and press Enter.');
    console.log('Example: https://www.tab.com.au/racing/2026-05-18/RANDWICK/NSW/R/1');
    console.log();

    const ask = () => {
      rl.question('TAB race URL: ', (answer) => {
        const url = answer.trim();
        if (url.startsWith('http')) {
          rl.close();
          resolve(url);
        } else {
          console.log('  Please enter a valid URL starting with http.');
          ask();
        }
      });
    };
    ask();
  });
}

// ---------------------------------------------------------------------------
// Output path helper
// ---------------------------------------------------------------------------

function resolveOutputPath(raw) {
  if (process.platform === 'win32') return raw;
  const lower = raw.toLowerCase();
  if (lower.startsWith('c:\\') || lower.startsWith('c:/')) {
    const rest = raw.slice(3).replace(/\\/g, '/');
    return path.join(os.homedir(), rest);
  }
  return raw.replace(/\\/g, '/');
}

// ---------------------------------------------------------------------------
// Browser helpers
// ---------------------------------------------------------------------------

async function buildBrowser() {
  const launcher = args.browser === 'firefox' ? firefox : chromium;
  return launcher.launch({
    headless: args.headless,
    args: ['--start-maximized', '--no-first-run', '--no-default-browser-check'],
  });
}

async function newPage(browser) {
  const ctx = await browser.newContext({
    userAgent:
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ' +
      'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    viewport: null,   // respect --start-maximized
  });
  const page = await ctx.newPage();
  page.setDefaultTimeout(15000);
  return page;
}

async function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

// ---------------------------------------------------------------------------
// TAB page — wait, cookie banner, Show All Form
// ---------------------------------------------------------------------------

async function waitForPage(page, timeout = 30000) {
  info('Waiting for runner list to appear...');
  await page.waitForSelector('.runner-name', { timeout });
  info('Page ready.');
}

async function dismissCookieBanner(page) {
  try {
    const btn = page.locator(
      "[class*='cookie'] button, [id*='cookie'] button, " +
      "[class*='consent'] button, [id*='consent'] button"
    ).first();
    if (await btn.count({ timeout: 3000 }) > 0) {
      await btn.click();
      await sleep(1000);
      info('Cookie banner dismissed.');
    }
  } catch { /* no banner — continue */ }
}

async function clickShowAllForm(page) {
  // Primary selector — matches the wrapper pattern used on TAB
  const primary = "[class*='show-all-form'] button, .show-all-form-wrapper button";
  try {
    const btn = page.locator(primary).first();
    await btn.waitFor({ state: 'visible', timeout: 10000 });
    const label = (await btn.innerText().catch(() => '')).trim();
    if (label.toLowerCase().includes('hide')) {
      info('Form already expanded — skipping click.');
      return true;
    }
    info(`Clicking "${label}" button...`);
    await btn.click();
    await page.waitForLoadState('networkidle', { timeout: 15000 }).catch(() => {});
    await sleep(4000);
    info('"Show All Form" expanded.');
    return true;
  } catch {
    warn('"Show All Form" button not found — trying individual expand buttons...');
    return expandIndividualForms(page);
  }
}

async function expandIndividualForms(page) {
  try {
    const btns = await page.locator(
      "[class*='form-toggle'] button, [class*='expand-form'] button, " +
      "[aria-label*='form' i], [aria-label*='Form']"
    ).all();
    for (const btn of btns) {
      const label = (await btn.getAttribute('aria-label').catch(() => '') || '').toLowerCase();
      const txt   = (await btn.innerText().catch(() => '')).toLowerCase();
      if (!txt.includes('hide') && !label.includes('hide')) {
        await btn.evaluate((el) => el.click());
        await sleep(300);
      }
    }
    await sleep(2000);
    info(`Expanded ${btns.length} individual form panels.`);
    return btns.length > 0;
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------------------
// Race info — parsed from URL + page headings
// ---------------------------------------------------------------------------

function parseRaceInfoFromUrl(url) {
  const info = { url };
  // e.g. /racing/2026-05-18/RANDWICK/NSW/R/3
  const m = url.match(/racing\/(\d{4}-\d{2}-\d{2})\/([A-Z0-9_-]+)\/[A-Z]+\/[RGH]\/(\d+)/i);
  if (m) {
    info.date     = m[1];
    info.venue    = m[2].replace(/-/g, ' ');
    info.raceNum  = m[3];
  }
  return info;
}

async function getRaceInfo(page, url) {
  const raceInfo = parseRaceInfoFromUrl(url);

  // Race name from page headings
  try {
    const headings = await page.locator('h1, h2, h3, h4').all();
    for (const h of headings) {
      const t = (await h.innerText().catch(() => '')).trim();
      if (t.length > 5 && !/TAB|Racing/i.test(t)) {
        raceInfo.raceName = t.replace(/\s*-\s*Betting Odds$/i, '').trim();
        break;
      }
    }
  } catch { raceInfo.raceName = ''; }

  // Banner details (distance, grade, prize money)
  try {
    const items = await page.locator(
      "banner li, [class*='race-info'] li, [class*='race-detail'] li"
    ).all();
    const texts = [];
    for (const el of items) {
      const t = (await el.innerText().catch(() => '')).trim();
      if (t) texts.push(t);
    }
    raceInfo.bannerDetails = texts.slice(0, 10).join(' | ');
  } catch { raceInfo.bannerDetails = ''; }

  // Track condition
  try {
    const conds = await page.locator("[class*='track-condition'], [class*='condition']").all();
    const parts = [];
    for (const c of conds) {
      const t = (await c.innerText().catch(() => '')).trim();
      if (t) parts.push(t);
    }
    raceInfo.trackCondition = parts.join(' ');
  } catch { raceInfo.trackCondition = ''; }

  return raceInfo;
}

// ---------------------------------------------------------------------------
// Runners — uses the real TAB CSS classes from greyhoundracing.py
// ---------------------------------------------------------------------------

async function scrapeRunners(page) {
  const runners = [];

  // All rows that contain a .runner-name element
  const allRows = await page.locator('.row').all();
  const runnerRows = [];
  for (const row of allRows) {
    if (await row.locator('.runner-name').count() > 0) runnerRows.push(row);
  }

  info(`Found ${runnerRows.length} runner rows`);

  for (const row of runnerRows) {
    const r = {
      number: '', name: '', jockey: '', trainer: '',
      form: '', weight: '', rating: '',
      winOdds: '', placeOdds: '', toteWin: '', totePlace: '',
      scratched: false,
    };

    // Runner number / barrier (horses use cloth number)
    try {
      r.number = (await row.locator(
        '.number-cell, [class*="number"], [class*="silk"]'
      ).first().innerText({ timeout: 1000 })).trim();
    } catch { /* fine */ }

    // Horse name — strip any nested span text
    try {
      const nameEl = row.locator('.runner-name').first();
      const nameRaw = await nameEl.innerText({ timeout: 2000 });
      const nested = await nameEl.locator('span, .barrier, .box').allInnerTexts();
      let name = nameRaw;
      for (const n of nested) name = name.replace(n, '');
      r.name = name.trim();
    } catch { /* skip row */ }

    if (!r.name) continue;

    // Jockey + Trainer from metadata list full-name elements
    try {
      const fullNames = await row.locator('.runner-metadata-list .full-name').allInnerTexts();
      r.jockey  = (fullNames[0] || '').trim();
      r.trainer = (fullNames[1] || '').trim();
    } catch { /* fine */ }

    // Form / Weight / Rating from optional metadata dt→dd pairs
    try {
      const dts = await row.locator('.runner-metadata-list.optional dt').allInnerTexts();
      const dds = await row.locator('.runner-metadata-list.optional dd').allInnerTexts();
      const meta = {};
      dts.forEach((k, i) => { meta[k.trim()] = (dds[i] || '').trim(); });
      r.form   = meta['F'] || meta['Form'] || '';
      r.weight = meta['W'] || meta['Weight'] || '';
      r.rating = meta['R'] || meta['Rating'] || '';
    } catch { /* fine */ }

    // Odds from price cells
    try {
      const prices = await row.locator('.price-cell').allInnerTexts();
      r.winOdds   = (prices[0] || '').trim();
      r.placeOdds = (prices[1] || '').trim();
      r.toteWin   = (prices[2] || '').trim();
      r.totePlace = (prices[3] || '').trim();
    } catch { /* fine */ }

    r.scratched = [r.winOdds, r.placeOdds].join(' ').includes('SCR');

    runners.push(r);
  }

  return runners;
}

// ---------------------------------------------------------------------------
// Form data — expanded panels after "Show All Form" click
// ---------------------------------------------------------------------------

function findVal(liTexts, key) {
  for (const t of liTexts) {
    if (t.startsWith(key)) return t.slice(key.length).trim();
  }
  return '';
}

async function scrapeFormData(page) {
  const formData = [];

  // Form rows sit after each runner row and carry class .row.form
  const formRows = await page.locator('.row.form').all();
  info(`Found ${formRows.length} expanded form panels`);

  for (const formRow of formRows) {
    let wrapper;
    try {
      wrapper = formRow.locator('.form-data-wrapper').first();
      await wrapper.waitFor({ state: 'attached', timeout: 3000 });
    } catch { continue; }

    const dog = { name: '', profile: {}, conditionStats: {}, raceHistory: [] };

    // Resolve the horse name from the preceding sibling runner row
    try {
      dog.name = await page.evaluate((el) => {
        let sib = el.previousElementSibling;
        while (sib) {
          const n = sib.querySelector('.runner-name');
          if (n) return n.textContent.trim();
          sib = sib.previousElementSibling;
        }
        return '';
      }, await formRow.elementHandle());
    } catch { dog.name = 'Unknown'; }

    // Profile fields from li elements inside the wrapper
    try {
      const liTexts = await wrapper.locator('li').allInnerTexts();
      const trimmed = liTexts.map((t) => t.trim()).filter(Boolean);

      const winsItem   = trimmed.find((t) => t.includes('Wins') && t.includes('%')) || '';
      const placesItem = trimmed.find((t) => t.includes('Places') && t.includes('%')) || '';
      const winsPct    = (winsItem.match(/(\d+)%/) || [])[1];
      const placesPct  = (placesItem.match(/(\d+)%/) || [])[1];

      dog.profile = {
        career:     findVal(trimmed, 'Career '),
        prizeMoney: findVal(trimmed, 'Prize Money '),
        sire:       findVal(trimmed, 'Sire '),
        dam:        findVal(trimmed, 'Dam '),
        colour:     findVal(trimmed, 'Colour '),
        sex:        findVal(trimmed, 'Sex '),
        age:        findVal(trimmed, 'Age '),
        trainer:    findVal(trimmed, 'Trainer '),
        jockey:     findVal(trimmed, 'Jockey '),
        owner:      findVal(trimmed, 'Owner '),
        lastRun:    findVal(trimmed, 'Last Run '),
        winsPct:    winsPct  ? winsPct  + '%' : '',
        placesPct:  placesPct ? placesPct + '%' : '',
      };

      dog.conditionStats = {
        track:    findVal(trimmed, 'Track '),
        distance: findVal(trimmed, 'Distance '),
        trkDist:  findVal(trimmed, 'Trk & Dist '),
        firm:     findVal(trimmed, 'Firm '),
        good:     findVal(trimmed, 'Good '),
        soft:     findVal(trimmed, 'Soft '),
        heavy:    findVal(trimmed, 'Heavy '),
        barrier:  findVal(trimmed, 'Barrier '),
        firstUp:  findVal(trimmed, '1st Up '),
        secondUp: findVal(trimmed, '2nd Up '),
        thirdUp:  findVal(trimmed, '3rd Up '),
      };
    } catch (e) {
      warn(`Error scraping profile for ${dog.name}: ${e.message}`);
    }

    // Race history from flexible row cells
    try {
      const bodyWrapper = wrapper.locator('.flexible-body-wrapper').first();
      const bodyRows = await bodyWrapper.locator('.flexible-row').all();

      for (const bRow of bodyRows) {
        // Spell / break row
        const spellEl = bRow.locator('.runner-spell .message');
        if (await spellEl.count() > 0) {
          dog.raceHistory.push({
            type: 'spell',
            message: (await spellEl.innerText().catch(() => '')).trim(),
          });
          continue;
        }

        const cells = await bRow.locator('.flexible-cell').allInnerTexts();
        const c = cells.map((s) => s.trim());
        if (c.length >= 8 && (c[0] || c[2])) {
          dog.raceHistory.push({
            type:      'race',
            placing:   c[0]  || '',
            venue:     c[1]  || '',
            date:      c[2]  || '',
            class:     c[3]  || '',
            distance:  c[4]  || '',
            weight:    c[5]  || '',
            barrier:   c[6]  || '',
            odds:      c[7]  || '',
            winner2nd: c[8]  || '',
            margin:    c[9]  || '',
            time:      c[10] || '',
            inRun:     c[11] || '',
          });
        }
      }
    } catch (e) {
      warn(`Error scraping race history for ${dog.name}: ${e.message}`);
    }

    formData.push(dog);
  }

  return formData;
}

// ---------------------------------------------------------------------------
// RacingZone scraper
// ---------------------------------------------------------------------------

async function scrapeRacingZoneHorse(page, horseName) {
  const stats = {
    name: horseName, sire: '', dam: '', colour: '', sex: '', age: '',
    country: '', owner: '', breeder: '',
    careerStarts: '', careerWins: '', careerSeconds: '', careerThirds: '',
    careerWinPct: '', careerPlacePct: '', careerPrizeMoney: '',
    l12mStarts: '', l12mWins: '', l12mSeconds: '', l12mThirds: '',
    statsByDistance: '', statsByCondition: '', statsByTrackType: '',
    statsByJockey: '', statsByTrainer: '',
    error: '',
  };

  info(`Searching RacingZone for: ${horseName}`);

  try {
    await page.goto(RACINGZONE_HORSES_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await sleep(2000);

    const inputSelectors = [
      "input[name='horse_name']", "input[name='name']",
      "input[placeholder*='horse' i]", "input[placeholder*='Find' i]",
      "input[type='search']", "#horse_name", "#name", "form input[type='text']",
    ];

    let searchInput = null;
    for (const sel of inputSelectors) {
      const el = page.locator(sel).first();
      if (await el.count() > 0) { searchInput = el; break; }
    }

    if (!searchInput) {
      stats.error = 'Search input not found';
      return stats;
    }

    await searchInput.fill(horseName);
    await sleep(500);

    const btnSelectors = [
      "button[type='submit']", "input[type='submit']",
      "button[class*='search' i]", "form button", "[class*='search-btn' i]",
    ];
    let submitted = false;
    for (const sel of btnSelectors) {
      const btn = page.locator(sel).first();
      if (await btn.count() > 0) { await btn.click(); submitted = true; break; }
    }
    if (!submitted) await searchInput.press('Enter');

    await sleep(3000);

    const resultLinks = await page.locator(
      '[class*="result"] a, table a, main a, #content a, .content a'
    ).all();

    if (resultLinks.length) {
      const nameLower = horseName.toLowerCase();
      let clicked = false;
      for (const link of resultLinks) {
        const txt = (await link.innerText().catch(() => '')).toLowerCase();
        if (txt.includes(nameLower) || nameLower.includes(txt)) {
          await link.click();
          clicked = true;
          await sleep(3000);
          break;
        }
      }
      if (!clicked) { await resultLinks[0].click(); await sleep(3000); }
    }

    await parseHorseStatsPage(page, stats);

  } catch (err) {
    error(`RacingZone error for ${horseName}:`, err.message);
    stats.error = err.message.slice(0, 200);
  }

  return stats;
}

async function parseHorseStatsPage(page, stats) {
  const pageText = await page.locator('body').innerText().catch(() => '');
  if (!pageText.trim()) { stats.error = 'Empty page'; return; }

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

  const tables = await page.locator('table').all();
  for (const tbl of tables) {
    for (const row of await tbl.locator('tr').all()) {
      const cells = await row.locator('td, th').all();
      if (cells.length < 2) continue;
      const label = (await cells[0].innerText().catch(() => '')).toLowerCase().replace(/:$/, '');
      const value = (await cells[1].innerText().catch(() => '')).trim();
      for (const [attr, kws] of Object.entries(detailMap)) {
        if (kws.some((kw) => label.includes(kw))) { stats[attr] = value; break; }
      }
    }
  }

  const dts = await page.locator('dt').allInnerTexts();
  const dds = await page.locator('dd').allInnerTexts();
  for (let i = 0; i < Math.min(dts.length, dds.length); i++) {
    const label = dts[i].toLowerCase().replace(/:$/, '');
    const value = dds[i].trim();
    for (const [attr, kws] of Object.entries(detailMap)) {
      if (kws.some((kw) => label.includes(kw))) { stats[attr] = value; break; }
    }
  }

  for (const tbl of tables) {
    const headers = await tbl.locator('th').allInnerTexts();
    const hdrs = headers.map((h) => h.toLowerCase());

    let dataRows = await tbl.locator('tbody tr').all();
    if (!dataRows.length) dataRows = (await tbl.locator('tr').all()).slice(1);

    for (const row of dataRows) {
      const cells = await row.locator('td').allInnerTexts();
      if (!cells.length) continue;
      const rowLabel = cells[0].toLowerCase();

      if (rowLabel.includes('career') || rowLabel.includes('total') || rowLabel.includes('all')) {
        fillStats(stats, cells, hdrs, 'career');
      } else if (rowLabel.includes('12') || rowLabel.includes('l12') || rowLabel.includes('year')) {
        fillStats(stats, cells, hdrs, 'l12m');
      }
    }
  }

  const sectionKeywords = {
    statsByDistance:  ['distance', 'dist'],
    statsByCondition: ['condition', 'going'],
    statsByTrackType: ['track type', 'surface'],
    statsByJockey:    ['jockey'],
    statsByTrainer:   ['trainer'],
  };
  for (const sec of await page.locator('section, [class*="section"], [class*="stats-group"]').all()) {
    const txt   = (await sec.innerText().catch(() => '')).trim();
    const lower = txt.toLowerCase();
    for (const [attr, kws] of Object.entries(sectionKeywords)) {
      if (!stats[attr] && kws.some((kw) => lower.includes(kw))) {
        stats[attr] = txt.slice(0, 2000);
      }
    }
  }
}

function fillStats(stats, cells, headers, prefix) {
  const mapping = {
    start: `${prefix}Starts`, win: `${prefix}Wins`,
    '2nd': `${prefix}Seconds`, second: `${prefix}Seconds`,
    '3rd': `${prefix}Thirds`,  third:  `${prefix}Thirds`,
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
// Excel export
// ---------------------------------------------------------------------------

const HEADER_COLOR = 'FF1B5E20';   // dark green (matches greyhoundracing.py)
const SECTION_COLOR = 'FF388E3C';  // mid green
const SPELL_COLOR = 'FFFFE699';    // yellow
const ALT_COLOR = 'FFE8F5E9';      // light green

function styleHeader(ws) {
  const row = ws.getRow(1);
  row.height = 22;
  row.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_COLOR } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = thinBorder();
  });
}

function thinBorder() {
  const s = { style: 'thin' };
  return { left: s, right: s, top: s, bottom: s };
}

function styleDataRow(ws, rowNum, alt = false) {
  const row = ws.getRow(rowNum);
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.fill = {
      type: 'pattern', pattern: 'solid',
      fgColor: { argb: alt ? ALT_COLOR : 'FFFFFFFF' },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    cell.border = thinBorder();
  });
}

function styleSectionRow(ws, rowNum, colCount) {
  for (let c = 1; c <= colCount; c++) {
    const cell = ws.getCell(rowNum, c);
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SECTION_COLOR } };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
    cell.border = thinBorder();
  }
}

function styleSpellRow(ws, rowNum, colCount) {
  for (let c = 1; c <= colCount; c++) {
    const cell = ws.getCell(rowNum, c);
    cell.font = { bold: true, italic: true, color: { argb: 'FF7F4F00' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SPELL_COLOR } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = thinBorder();
  }
}

function autoWidth(ws, min = 8, max = 40) {
  ws.columns.forEach((col) => {
    let maxLen = 0;
    col.eachCell({ includeEmpty: true }, (cell) => {
      const v = cell.value ? String(cell.value) : '';
      maxLen = Math.max(maxLen, v.length);
    });
    col.width = Math.min(Math.max(maxLen + 2, min), max);
  });
}

async function saveToExcel(raceInfo, runners, formData, horseStats, outputPath) {
  fs.mkdirSync(outputPath, { recursive: true });

  const safeVenue = (raceInfo.venue || 'Unknown').replace(/[^a-zA-Z0-9 _-]/g, '_');
  const dateStr   = (raceInfo.date || new Date().toISOString().slice(0, 10)).replace(/-/g, '');
  const filename  = `TAB_${safeVenue}_R${raceInfo.raceNum || '0'}_${dateStr}.xlsx`.replace(/\s+/g, '_');
  const filepath  = path.join(outputPath, filename);

  info('Saving Excel:', filepath);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'RacingZoneScraper';
  wb.created = new Date();

  // ── Sheet 1: Race Info ────────────────────────────────────────────────────
  const ws1 = wb.addWorksheet('Race Info');
  ws1.columns = [
    { header: 'Field', key: 'field', width: 20 },
    { header: 'Value', key: 'value', width: 50 },
  ];
  styleHeader(ws1);
  ws1.views = [{ state: 'frozen', ySplit: 1 }];

  const infoRows = [
    ['Race Name',       raceInfo.raceName || ''],
    ['Date',            raceInfo.date     || ''],
    ['Venue',           raceInfo.venue    || ''],
    ['Race Number',     raceInfo.raceNum  || ''],
    ['Track Condition', raceInfo.trackCondition || ''],
    ['Banner Details',  raceInfo.bannerDetails  || ''],
    ['URL',             raceInfo.url || ''],
    ['Scraped At',      new Date().toLocaleString()],
  ];
  infoRows.forEach(([field, value], i) => {
    ws1.addRow({ field, value });
    styleDataRow(ws1, i + 2, i % 2 === 1);
  });
  autoWidth(ws1);

  // ── Sheet 2: Runners (TAB) ────────────────────────────────────────────────
  const ws2 = wb.addWorksheet('Runners (TAB)');
  ws2.columns = [
    { header: 'No.',        key: 'number',    width: 8  },
    { header: 'Horse',      key: 'name',      width: 25 },
    { header: 'Jockey',     key: 'jockey',    width: 22 },
    { header: 'Trainer',    key: 'trainer',   width: 22 },
    { header: 'Form',       key: 'form',      width: 14 },
    { header: 'Weight',     key: 'weight',    width: 10 },
    { header: 'Rating',     key: 'rating',    width: 10 },
    { header: 'FO Win',     key: 'winOdds',   width: 10 },
    { header: 'FO Place',   key: 'placeOdds', width: 10 },
    { header: 'Tote Win',   key: 'toteWin',   width: 10 },
    { header: 'Tote Place', key: 'totePlace', width: 10 },
    { header: 'Scratched',  key: 'scratched', width: 10 },
  ];
  styleHeader(ws2);
  ws2.views = [{ state: 'frozen', ySplit: 1 }];

  runners.forEach((r, i) => {
    ws2.addRow({ ...r, scratched: r.scratched ? 'Yes' : 'No' });
    styleDataRow(ws2, i + 2, i % 2 === 1);
    if (r.scratched) {
      ws2.getRow(i + 2).eachCell((cell) => {
        cell.font = { italic: true, color: { argb: 'FF999999' } };
      });
    }
  });
  autoWidth(ws2);

  // ── Sheet 3: Horse Profiles ───────────────────────────────────────────────
  const ws3 = wb.addWorksheet('Horse Profiles');
  ws3.columns = [
    { header: 'No.',         key: 'number',    width: 8  },
    { header: 'Horse',       key: 'name',      width: 25 },
    { header: 'Career',      key: 'career',    width: 18 },
    { header: 'Prize Money', key: 'prize',     width: 14 },
    { header: 'Sire',        key: 'sire',      width: 20 },
    { header: 'Dam',         key: 'dam',       width: 20 },
    { header: 'Colour',      key: 'colour',    width: 12 },
    { header: 'Sex',         key: 'sex',       width: 8  },
    { header: 'Age',         key: 'age',       width: 8  },
    { header: 'Owner',       key: 'owner',     width: 22 },
    { header: 'Trainer',     key: 'trainer',   width: 22 },
    { header: 'Jockey',      key: 'jockey',    width: 22 },
    { header: 'Last Run',    key: 'lastRun',   width: 14 },
    { header: 'Wins %',      key: 'winsPct',   width: 10 },
    { header: 'Places %',    key: 'placesPct', width: 10 },
  ];
  styleHeader(ws3);
  ws3.views = [{ state: 'frozen', ySplit: 1 }];

  const runnerMap = new Map(runners.map((r) => [r.name.toUpperCase(), r]));

  formData.forEach((fd, i) => {
    const cleanName = fd.name.replace(/\s*\(\d+\)\s*$/, '').trim();
    const runner = runnerMap.get(cleanName.toUpperCase()) || {};
    const p = fd.profile;
    ws3.addRow({
      number: runner.number || '', name: cleanName,
      career: p.career, prize: p.prizeMoney,
      sire: p.sire, dam: p.dam, colour: p.colour, sex: p.sex, age: p.age,
      owner: p.owner, trainer: p.trainer || runner.trainer,
      jockey: p.jockey || runner.jockey,
      lastRun: p.lastRun, winsPct: p.winsPct, placesPct: p.placesPct,
    });
    styleDataRow(ws3, i + 2, i % 2 === 1);
  });
  autoWidth(ws3);

  // ── Sheet 4: Condition Stats ──────────────────────────────────────────────
  const ws4 = wb.addWorksheet('Condition Stats');
  ws4.columns = [
    { header: 'No.',       key: 'number',   width: 8  },
    { header: 'Horse',     key: 'name',     width: 25 },
    { header: 'Track',     key: 'track',    width: 16 },
    { header: 'Distance',  key: 'distance', width: 16 },
    { header: 'Trk&Dist',  key: 'trkDist',  width: 16 },
    { header: 'Firm',      key: 'firm',     width: 14 },
    { header: 'Good',      key: 'good',     width: 14 },
    { header: 'Soft',      key: 'soft',     width: 14 },
    { header: 'Heavy',     key: 'heavy',    width: 14 },
    { header: 'Barrier',   key: 'barrier',  width: 14 },
    { header: '1st Up',    key: 'firstUp',  width: 14 },
    { header: '2nd Up',    key: 'secondUp', width: 14 },
    { header: '3rd Up',    key: 'thirdUp',  width: 14 },
  ];
  styleHeader(ws4);
  ws4.views = [{ state: 'frozen', ySplit: 1 }];

  formData.forEach((fd, i) => {
    const cleanName = fd.name.replace(/\s*\(\d+\)\s*$/, '').trim();
    const runner = runnerMap.get(cleanName.toUpperCase()) || {};
    const cs = fd.conditionStats;
    ws4.addRow({
      number: runner.number || '', name: cleanName,
      track: cs.track, distance: cs.distance, trkDist: cs.trkDist,
      firm: cs.firm, good: cs.good, soft: cs.soft, heavy: cs.heavy,
      barrier: cs.barrier, firstUp: cs.firstUp, secondUp: cs.secondUp, thirdUp: cs.thirdUp,
    });
    styleDataRow(ws4, i + 2, i % 2 === 1);
  });
  autoWidth(ws4);

  // ── Sheet 5: Race History ─────────────────────────────────────────────────
  const ws5 = wb.addWorksheet('Race History');
  const histCols = [
    { header: 'No.',       key: 'number',   width: 8  },
    { header: 'Horse',     key: 'name',     width: 25 },
    { header: 'Placing',   key: 'placing',  width: 10 },
    { header: 'Venue',     key: 'venue',    width: 18 },
    { header: 'Date',      key: 'date',     width: 12 },
    { header: 'Class',     key: 'class',    width: 16 },
    { header: 'Dist',      key: 'distance', width: 10 },
    { header: 'Weight',    key: 'weight',   width: 10 },
    { header: 'Barrier',   key: 'barrier',  width: 10 },
    { header: 'Odds',      key: 'odds',     width: 10 },
    { header: 'Winner/2nd',key: 'winner2nd',width: 22 },
    { header: 'Margin',    key: 'margin',   width: 10 },
    { header: 'Time',      key: 'time',     width: 10 },
    { header: 'In Run',    key: 'inRun',    width: 14 },
  ];
  ws5.columns = histCols;
  styleHeader(ws5);
  ws5.views = [{ state: 'frozen', xSplit: 2, ySplit: 1 }];

  let currentRow = 2;
  for (const fd of formData) {
    const cleanName = fd.name.replace(/\s*\(\d+\)\s*$/, '').trim();
    const runner = runnerMap.get(cleanName.toUpperCase()) || {};
    const boxNum = runner.number || '';

    // Section header row per horse
    ws5.getCell(currentRow, 1).value = boxNum;
    ws5.getCell(currentRow, 2).value = cleanName;
    for (let c = 3; c <= histCols.length; c++) ws5.getCell(currentRow, c).value = '';
    styleSectionRow(ws5, currentRow, histCols.length);
    currentRow++;

    let alt = false;
    for (const entry of fd.raceHistory) {
      if (entry.type === 'spell') {
        ws5.getCell(currentRow, 1).value = entry.message;
        ws5.mergeCells(currentRow, 1, currentRow, histCols.length);
        styleSpellRow(ws5, currentRow, histCols.length);
        currentRow++;
        alt = false;
      } else {
        ws5.addRow({
          number: boxNum, name: cleanName,
          placing: entry.placing, venue: entry.venue, date: entry.date,
          class: entry.class, distance: entry.distance, weight: entry.weight,
          barrier: entry.barrier, odds: entry.odds, winner2nd: entry.winner2nd,
          margin: entry.margin, time: entry.time, inRun: entry.inRun,
        });
        styleDataRow(ws5, currentRow, alt);
        currentRow++;
        alt = !alt;
      }
    }
  }
  autoWidth(ws5);

  // ── Sheet 6: RacingZone Stats ─────────────────────────────────────────────
  const ws6 = wb.addWorksheet('RacingZone Stats');
  ws6.columns = [
    { header: 'Horse',           key: 'name',            width: 25 },
    { header: 'Career Starts',   key: 'careerStarts',    width: 14 },
    { header: 'Career Wins',     key: 'careerWins',      width: 13 },
    { header: 'Career 2nds',     key: 'careerSeconds',   width: 13 },
    { header: 'Career 3rds',     key: 'careerThirds',    width: 13 },
    { header: 'Career Win %',    key: 'careerWinPct',    width: 13 },
    { header: 'Career Place %',  key: 'careerPlacePct',  width: 14 },
    { header: 'Career Prize $',  key: 'careerPrizeMoney',width: 16 },
    { header: 'L12M Starts',     key: 'l12mStarts',      width: 13 },
    { header: 'L12M Wins',       key: 'l12mWins',        width: 12 },
    { header: 'L12M 2nds',       key: 'l12mSeconds',     width: 12 },
    { header: 'L12M 3rds',       key: 'l12mThirds',      width: 12 },
    { header: 'By Distance',     key: 'statsByDistance', width: 40 },
    { header: 'By Condition',    key: 'statsByCondition',width: 40 },
    { header: 'By Track Type',   key: 'statsByTrackType',width: 40 },
    { header: 'By Jockey',       key: 'statsByJockey',   width: 40 },
    { header: 'By Trainer',      key: 'statsByTrainer',  width: 40 },
    { header: 'Error',           key: 'error',           width: 30 },
  ];
  styleHeader(ws6);
  ws6.views = [{ state: 'frozen', ySplit: 1 }];
  horseStats.forEach((s, i) => {
    ws6.addRow(s);
    styleDataRow(ws6, i + 2, i % 2 === 1);
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
  const url = await promptForUrl();
  const outputPath = resolveOutputPath(args.output);
  info('Output directory:', outputPath);

  info('Launching Chromium...');
  const browser = await buildBrowser();
  const page = await newPage(browser);

  try {
    // Step 1: Load the TAB race page
    info('Navigating to:', url);
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });

    await dismissCookieBanner(page);
    await waitForPage(page);
    await sleep(2000);

    // Step 2: Click "Show All Form" to expand all form panels
    info('Clicking "Show All Form"...');
    await clickShowAllForm(page);
    await sleep(3000);

    // Step 3: Scrape race metadata and runners
    info('Scraping race info...');
    const raceInfo = await getRaceInfo(page, url);
    info(`Race: ${raceInfo.raceName} | ${raceInfo.date} | ${raceInfo.venue} | R${raceInfo.raceNum}`);

    info('Scraping runners...');
    const runners = await scrapeRunners(page);
    const active    = runners.filter((r) => !r.scratched);
    const scratched = runners.filter((r) => r.scratched);
    info(`Runners: ${runners.length} total (${active.length} active, ${scratched.length} scratched)`);

    if (!runners.length) {
      error('No runners found. Check the URL and try again. Page title:', await page.title());
      process.exit(1);
    }

    // Step 4: Scrape expanded form data
    info('Scraping form data...');
    const formData = await scrapeFormData(page);
    info(`Form data scraped for ${formData.length} horses`);

    // Step 5: RacingZone lookup for each active horse
    info('Starting RacingZone lookups...');
    const horseStats = [];
    for (let i = 0; i < runners.length; i++) {
      const runner = runners[i];
      if (runner.scratched) {
        horseStats.push({ name: runner.name, error: 'Scratched' });
        continue;
      }
      info(`[${i + 1}/${runners.length}] RacingZone: ${runner.name}`);
      const stats = await scrapeRacingZoneHorse(page, runner.name);
      horseStats.push(stats);
      if (i < runners.length - 1) await sleep(args.delayMs);
    }

    // Step 6: Save to Excel
    const outputFile = await saveToExcel(raceInfo, runners, formData, horseStats, outputPath);
    console.log(`\nDone! Excel file saved to: ${outputFile}`);

  } finally {
    await browser.close();
  }
}

main().catch((err) => {
  error('Fatal error:', err.message);
  process.exit(1);
});
