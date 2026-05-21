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
    headless: true,
    delayMs: DEFAULT_DELAY_MS,
  };
  for (let i = 0; i < argv.length; i++) {
    switch (argv[i]) {
      case '--output':   opts.output   = argv[++i]; break;
      case '--browser':  opts.browser  = argv[++i]; break;
      case '--no-headless': opts.headless = false;  break;
      case '--delay':    opts.delayMs  = Number(argv[++i]) * 1000; break;
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
    console.log('Example: https://www.tab.com.au/racing/meetings/RANDWICK/...');
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
  // Map C:\tab\scrape → ~/tab/scrape on Linux/Mac
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
  const browser = await launcher.launch({
    headless: args.headless,
    args: ['--no-sandbox', '--disable-dev-shm-usage'],
  });
  return browser;
}

async function newPage(browser) {
  const ctx = await browser.newContext({
    userAgent:
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ' +
      'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    viewport: { width: 1920, height: 1080 },
  });
  const page = await ctx.newPage();
  page.setDefaultTimeout(15000);
  return page;
}

async function safeText(locator) {
  try { return (await locator.innerText()).trim(); } catch { return ''; }
}

async function firstText(page, selectors) {
  for (const sel of selectors) {
    try {
      const el = page.locator(sel).first();
      const txt = (await el.innerText({ timeout: 2000 })).trim();
      if (txt) return txt;
    } catch { /* try next */ }
  }
  return '';
}

async function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

// ---------------------------------------------------------------------------
// TAB scraper
// ---------------------------------------------------------------------------

async function scrapeTabPage(page, url) {
  info('Navigating to TAB page:', url);
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await sleep(3000);

  const race = {
    venue: '', raceNumber: '', raceName: '', date: '', time: '',
    distance: '', trackCondition: '', prizeMoney: '', raceClass: '',
    runners: [],
  };

  // --- Race metadata ---
  race.venue = await firstText(page, [
    '[class*="meeting-name"]', '[class*="MeetingName"]',
    '[class*="venue"]', 'h1',
  ]);
  race.raceNumber = await firstText(page, [
    '[class*="race-number"]', '[class*="RaceNumber"]', '[class*="raceNumber"]',
  ]);
  race.raceName = await firstText(page, [
    '[class*="race-name"]', '[class*="RaceName"]', '[class*="raceName"]', 'h2',
  ]);
  race.distance = await firstText(page, [
    '[class*="distance"]', '[class*="Distance"]',
  ]);
  race.trackCondition = await firstText(page, [
    '[class*="track-condition"]', '[class*="TrackCondition"]', '[class*="condition"]',
  ]);
  race.prizeMoney = await firstText(page, [
    '[class*="prize-money"]', '[class*="PrizeMoney"]', '[class*="prize"]',
  ]);
  race.date = await firstText(page, ['[class*="date"]', 'time']);

  info(`Race metadata — venue: ${race.venue}, race: ${race.raceNumber}, name: ${race.raceName}, dist: ${race.distance}`);

  // --- Runners ---
  race.runners = await scrapeRunners(page);
  info(`Found ${race.runners.length} runners`);

  return race;
}

async function scrapeRunners(page) {
  const rowSelectors = [
    '[class*="runner-row"]', '[class*="RunnerRow"]', '[class*="runner_row"]',
    'tr[class*="runner"]', '[class*="race-runner"]', '[data-testid*="runner"]',
    '[class*="competitor"]', 'tbody tr',
  ];

  let rowHandles = [];
  let usedSel = '';
  for (const sel of rowSelectors) {
    rowHandles = await page.locator(sel).all();
    if (rowHandles.length) { usedSel = sel; break; }
  }

  if (!rowHandles.length) {
    warn('No runner rows found — falling back to text parse');
    return parseRunnersFromText(page);
  }

  info(`Runner rows found with selector: ${usedSel} (${rowHandles.length} rows)`);

  const runners = [];
  for (const row of rowHandles) {
    const runner = {
      number: '', name: '', barrier: '', jockey: '', trainer: '',
      weight: '', age: '', sex: '', form: '', winOdds: '', placeOdds: '',
      scratched: false,
    };

    runner.number = await firstTextIn(row, [
      '[class*="number"]', '[class*="Number"]', 'td:first-child', '[class*="silk"]',
    ]);
    runner.name = await firstTextIn(row, [
      '[class*="horse-name"]', '[class*="HorseName"]',
      '[class*="runner-name"]', '[class*="RunnerName"]',
      '[class*="name"]', 'a',
    ]);
    runner.barrier = await firstTextIn(row, [
      '[class*="barrier"]', '[class*="Barrier"]', '[class*="gate"]',
    ]);
    runner.jockey = await firstTextIn(row, [
      '[class*="jockey"]', '[class*="Jockey"]', '[class*="rider"]',
    ]);
    runner.trainer = await firstTextIn(row, [
      '[class*="trainer"]', '[class*="Trainer"]',
    ]);
    runner.weight = await firstTextIn(row, [
      '[class*="weight"]', '[class*="Weight"]', '[class*="carried"]',
    ]);
    runner.form = await firstTextIn(row, [
      '[class*="form"]', '[class*="Form"]',
    ]);
    runner.winOdds = await firstTextIn(row, [
      '[class*="win-price"]', '[class*="WinPrice"]',
      '[class*="win-odds"]', '[class*="WinOdds"]', '[class*="fixed-win"]',
    ]);
    runner.placeOdds = await firstTextIn(row, [
      '[class*="place-price"]', '[class*="PlacePrice"]',
      '[class*="place-odds"]', '[class*="PlaceOdds"]', '[class*="fixed-place"]',
    ]);

    const scratchEl = row.locator('[class*="scratch"]');
    runner.scratched = (await scratchEl.count()) > 0;

    if (runner.name) runners.push(runner);
    else {
      const raw = (await row.innerText().catch(() => '')).trim();
      if (raw) debug('Skipping row with no name — raw:', raw.slice(0, 100));
    }
  }

  return runners;
}

async function firstTextIn(rowLocator, selectors) {
  for (const sel of selectors) {
    try {
      const el = rowLocator.locator(sel).first();
      const txt = (await el.innerText({ timeout: 1000 })).trim();
      if (txt) return txt;
    } catch { /* try next */ }
  }
  return '';
}

async function parseRunnersFromText(page) {
  const runners = [];
  const body = await page.locator('body').innerText().catch(() => '');
  const lines = body.split('\n').map((l) => l.trim()).filter(Boolean);
  const pattern = /^(\d{1,2})\s+([A-Z][A-Z\s'()-]+)$/;
  for (const line of lines) {
    const m = line.match(pattern);
    if (m) runners.push({ number: m[1], name: m[2].trim(), barrier: '', jockey: '',
      trainer: '', weight: '', age: '', sex: '', form: '', winOdds: '', placeOdds: '',
      scratched: false });
  }
  return runners;
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

    // Find search input
    const inputSelectors = [
      "input[name='horse_name']", "input[name='name']",
      "input[placeholder*='horse' i]", "input[placeholder*='Find' i]",
      "input[type='search']", "#horse_name", "#name", "form input[type='text']",
    ];

    let searchInput = null;
    for (const sel of inputSelectors) {
      const el = page.locator(sel).first();
      if (await el.count() > 0) { searchInput = el; debug('Search input:', sel); break; }
    }

    if (!searchInput) {
      warn(`Could not find search input for ${horseName}`);
      stats.error = 'Search input not found';
      return stats;
    }

    await searchInput.fill(horseName);
    await sleep(500);

    // Submit
    const btnSelectors = [
      "button[type='submit']", "input[type='submit']",
      "button[class*='search' i]", "form button", "[class*='search-btn' i]",
    ];
    let submitted = false;
    for (const sel of btnSelectors) {
      const btn = page.locator(sel).first();
      if (await btn.count() > 0) {
        await btn.click();
        submitted = true;
        debug('Search button clicked:', sel);
        break;
      }
    }
    if (!submitted) await searchInput.press('Enter');

    await sleep(3000);

    // If results list appeared, click best matching link
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
    error(`Error fetching RacingZone for ${horseName}:`, err.message);
    stats.error = err.message.slice(0, 200);
  }

  info(`Parsed stats for ${horseName} — career: ${stats.careerStarts}/${stats.careerWins}/${stats.careerSeconds}/${stats.careerThirds}`);
  return stats;
}

async function parseHorseStatsPage(page, stats) {
  const pageText = await page.locator('body').innerText().catch(() => '');
  if (!pageText.trim()) { stats.error = 'Empty page'; return; }

  // Horse name from heading
  for (const sel of ['[class*="horse-name"]', '[class*="HorseName"]', 'h1', 'h2']) {
    try {
      const txt = (await page.locator(sel).first().innerText({ timeout: 2000 })).trim();
      if (txt) { stats.name = txt; break; }
    } catch { /* skip */ }
  }

  // Profile fields via tables (label | value rows)
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
    const rows = await tbl.locator('tr').all();
    for (const row of rows) {
      const cells = await row.locator('td, th').all();
      if (cells.length < 2) continue;
      const label = (await cells[0].innerText().catch(() => '')).toLowerCase().replace(/:$/, '');
      const value = (await cells[1].innerText().catch(() => '')).trim();
      for (const [attr, keywords] of Object.entries(detailMap)) {
        if (keywords.some((kw) => label.includes(kw))) { stats[attr] = value; break; }
      }
    }
  }

  // dl/dt/dd
  const dts = await page.locator('dt').all();
  const dds = await page.locator('dd').all();
  for (let i = 0; i < Math.min(dts.length, dds.length); i++) {
    const label = (await dts[i].innerText().catch(() => '')).toLowerCase().replace(/:$/, '');
    const value = (await dds[i].innerText().catch(() => '')).trim();
    for (const [attr, keywords] of Object.entries(detailMap)) {
      if (keywords.some((kw) => label.includes(kw))) { stats[attr] = value; break; }
    }
  }

  // Career / L12M stats from tables
  for (const tbl of tables) {
    const headers = [];
    for (const th of await tbl.locator('th').all()) {
      headers.push((await th.innerText().catch(() => '')).toLowerCase());
    }

    let dataRows = await tbl.locator('tbody tr').all();
    if (!dataRows.length) {
      const all = await tbl.locator('tr').all();
      dataRows = all.slice(1);
    }

    for (const row of dataRows) {
      const cells = [];
      for (const td of await row.locator('td').all()) {
        cells.push((await td.innerText().catch(() => '')).trim());
      }
      if (!cells.length) continue;
      const rowLabel = cells[0].toLowerCase();

      if (rowLabel.includes('career') || rowLabel.includes('total') || rowLabel.includes('all')) {
        fillStats(stats, cells, headers, 'career');
      } else if (rowLabel.includes('12') || rowLabel.includes('l12') || rowLabel.includes('year')) {
        fillStats(stats, cells, headers, 'l12m');
      }
    }
  }

  // Stats-by-X sections as raw text
  const sectionSelectors = 'section, [class*="section"], [class*="stats-group"], [class*="StatsGroup"]';
  const sections = await page.locator(sectionSelectors).all();
  const sectionKeywords = {
    statsByDistance:  ['distance', 'dist'],
    statsByCondition: ['condition', 'going', 'track cond'],
    statsByTrackType: ['track type', 'surface'],
    statsByJockey:    ['jockey'],
    statsByTrainer:   ['trainer'],
  };
  for (const sec of sections) {
    const txt = (await sec.innerText().catch(() => '')).trim();
    const lower = txt.toLowerCase();
    for (const [attr, keywords] of Object.entries(sectionKeywords)) {
      if (!stats[attr] && keywords.some((kw) => lower.includes(kw))) {
        stats[attr] = txt.slice(0, 2000);
      }
    }
  }

  // Fallback: extract from plain page text
  if (!stats.statsByDistance && !stats.statsByCondition) {
    extractSectionsFromText(stats, pageText);
  }
}

function fillStats(stats, cells, headers, prefix) {
  const mapping = {
    start:    `${prefix}Starts`,
    win:      `${prefix}Wins`,
    '2nd':    `${prefix}Seconds`,
    second:   `${prefix}Seconds`,
    '3rd':    `${prefix}Thirds`,
    third:    `${prefix}Thirds`,
    'win%':   `${prefix}WinPct`,
    'place%': `${prefix}PlacePct`,
    prize:    `${prefix}PrizeMoney`,
    earning:  `${prefix}PrizeMoney`,
  };
  cells.slice(1).forEach((val, idx) => {
    const hdr = (headers[idx] || '').toLowerCase();
    for (const [key, attr] of Object.entries(mapping)) {
      if (hdr.includes(key) && attr in stats) { stats[attr] = val; break; }
    }
  });
}

function extractSectionsFromText(stats, pageText) {
  const lines = pageText.split('\n');
  const sections = { distance: [], condition: [], jockey: [], trainer: [] };
  let current = null;
  for (const line of lines) {
    const ll = line.toLowerCase().trim();
    if (ll.includes('distance'))           current = 'distance';
    else if (ll.includes('condition') || ll.includes('going')) current = 'condition';
    else if (ll.includes('jockey'))        current = 'jockey';
    else if (ll.includes('trainer'))       current = 'trainer';
    else if (current) sections[current].push(line);
  }
  if (sections.distance.length)  stats.statsByDistance  = sections.distance.slice(0, 20).join('\n');
  if (sections.condition.length) stats.statsByCondition = sections.condition.slice(0, 20).join('\n');
  if (sections.jockey.length)    stats.statsByJockey    = sections.jockey.slice(0, 20).join('\n');
  if (sections.trainer.length)   stats.statsByTrainer   = sections.trainer.slice(0, 20).join('\n');
}

// ---------------------------------------------------------------------------
// Excel export
// ---------------------------------------------------------------------------

async function saveToExcel(race, horseStats, outputPath) {
  fs.mkdirSync(outputPath, { recursive: true });

  const safeVenue = (race.venue || 'Unknown').replace(/[^a-zA-Z0-9 _-]/g, '_');
  const filename = `${safeVenue}_R${race.raceNumber || '0'}.xlsx`.replace(/\s+/g, '_');
  const filepath = path.join(outputPath, filename);

  info('Saving Excel file:', filepath);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'RacingZoneScraper';
  wb.created = new Date();

  // --- Sheet 1: Race Info ---
  const sheetRace = wb.addWorksheet('Race Info');
  sheetRace.columns = [
    { header: 'Field', key: 'field', width: 20 },
    { header: 'Value', key: 'value', width: 40 },
  ];
  styleHeader(sheetRace);
  const raceFields = [
    ['Venue', race.venue], ['Race Number', race.raceNumber],
    ['Race Name', race.raceName], ['Date', race.date], ['Time', race.time],
    ['Distance', race.distance], ['Track Condition', race.trackCondition],
    ['Prize Money', race.prizeMoney], ['Race Class', race.raceClass],
  ];
  raceFields.forEach(([field, value]) => sheetRace.addRow({ field, value }));

  // --- Sheet 2: Runners ---
  const sheetRunners = wb.addWorksheet('Runners (TAB)');
  sheetRunners.columns = [
    { header: 'Number',      key: 'number',     width: 10 },
    { header: 'Horse Name',  key: 'name',        width: 25 },
    { header: 'Barrier',     key: 'barrier',     width: 10 },
    { header: 'Jockey',      key: 'jockey',      width: 22 },
    { header: 'Trainer',     key: 'trainer',     width: 22 },
    { header: 'Weight (kg)', key: 'weight',      width: 12 },
    { header: 'Age',         key: 'age',         width: 8  },
    { header: 'Sex',         key: 'sex',         width: 8  },
    { header: 'Form',        key: 'form',        width: 14 },
    { header: 'Win Odds',    key: 'winOdds',     width: 12 },
    { header: 'Place Odds',  key: 'placeOdds',   width: 12 },
    { header: 'Scratched',   key: 'scratched',   width: 12 },
  ];
  styleHeader(sheetRunners);
  race.runners.forEach((r) =>
    sheetRunners.addRow({ ...r, scratched: r.scratched ? 'Yes' : 'No' })
  );

  // --- Sheet 3: Horse Stats ---
  const sheetStats = wb.addWorksheet('Horse Stats (RacingZone)');
  sheetStats.columns = [
    { header: 'Horse Name',          key: 'name',             width: 25 },
    { header: 'Sire',                key: 'sire',             width: 20 },
    { header: 'Dam',                 key: 'dam',              width: 20 },
    { header: 'Colour',              key: 'colour',           width: 14 },
    { header: 'Sex',                 key: 'sex',              width: 8  },
    { header: 'Age',                 key: 'age',              width: 8  },
    { header: 'Country',             key: 'country',          width: 12 },
    { header: 'Owner',               key: 'owner',            width: 22 },
    { header: 'Breeder',             key: 'breeder',          width: 22 },
    { header: 'Career Starts',       key: 'careerStarts',     width: 14 },
    { header: 'Career Wins',         key: 'careerWins',       width: 13 },
    { header: 'Career 2nds',         key: 'careerSeconds',    width: 13 },
    { header: 'Career 3rds',         key: 'careerThirds',     width: 13 },
    { header: 'Career Win %',        key: 'careerWinPct',     width: 13 },
    { header: 'Career Place %',      key: 'careerPlacePct',   width: 14 },
    { header: 'Career Prize Money',  key: 'careerPrizeMoney', width: 18 },
    { header: 'L12M Starts',         key: 'l12mStarts',       width: 13 },
    { header: 'L12M Wins',           key: 'l12mWins',         width: 12 },
    { header: 'L12M 2nds',           key: 'l12mSeconds',      width: 12 },
    { header: 'L12M 3rds',           key: 'l12mThirds',       width: 12 },
    { header: 'Stats by Distance',   key: 'statsByDistance',  width: 40 },
    { header: 'Stats by Condition',  key: 'statsByCondition', width: 40 },
    { header: 'Stats by Track Type', key: 'statsByTrackType', width: 40 },
    { header: 'Stats by Jockey',     key: 'statsByJockey',    width: 40 },
    { header: 'Stats by Trainer',    key: 'statsByTrainer',   width: 40 },
    { header: 'Error',               key: 'error',            width: 30 },
  ];
  styleHeader(sheetStats);
  horseStats.forEach((s) => sheetStats.addRow(s));

  // --- Sheet 4: Combined ---
  const sheetCombined = wb.addWorksheet('Combined');
  const runnerMap = new Map(race.runners.map((r) => [r.name, r]));
  const combinedCols = [
    ...sheetRunners.columns.map((c) => ({ ...c })),
    ...sheetStats.columns.slice(1).map((c) => ({ ...c })),
  ];
  sheetCombined.columns = combinedCols;
  styleHeader(sheetCombined);
  horseStats.forEach((s) => {
    const r = runnerMap.get(s.name) || {};
    sheetCombined.addRow({
      number: r.number || '', name: s.name, barrier: r.barrier || '',
      jockey: r.jockey || '', trainer: r.trainer || '', weight: r.weight || '',
      age: r.age || s.age, sex: r.sex || s.sex, form: r.form || '',
      winOdds: r.winOdds || '', placeOdds: r.placeOdds || '',
      scratched: r.scratched ? 'Yes' : 'No',
      sire: s.sire, dam: s.dam, colour: s.colour, country: s.country,
      owner: s.owner, breeder: s.breeder,
      careerStarts: s.careerStarts, careerWins: s.careerWins,
      careerSeconds: s.careerSeconds, careerThirds: s.careerThirds,
      careerWinPct: s.careerWinPct, careerPlacePct: s.careerPlacePct,
      careerPrizeMoney: s.careerPrizeMoney,
      l12mStarts: s.l12mStarts, l12mWins: s.l12mWins,
      l12mSeconds: s.l12mSeconds, l12mThirds: s.l12mThirds,
      statsByDistance: s.statsByDistance, statsByCondition: s.statsByCondition,
      statsByTrackType: s.statsByTrackType, statsByJockey: s.statsByJockey,
      statsByTrainer: s.statsByTrainer, error: s.error,
    });
  });

  await wb.xlsx.writeFile(filepath);
  info('Excel saved:', filepath);
  return filepath;
}

function styleHeader(ws) {
  const headerRow = ws.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
  headerRow.height = 20;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  const url = await promptForUrl();
  const outputPath = resolveOutputPath(args.output);
  info('Output directory:', outputPath);

  const browser = await buildBrowser();
  const page = await newPage(browser);

  try {
    // Step 1: Scrape TAB race page
    const race = await scrapeTabPage(page, url);

    if (!race.runners.length) {
      error('No runners found on TAB page. Check the URL and try again.');
      error('Page title:', await page.title());
      process.exit(1);
    }

    info(`Found ${race.runners.length} runners. Starting RacingZone lookups...`);

    // Step 2: Scrape RacingZone for each horse
    const horseStats = [];
    for (let i = 0; i < race.runners.length; i++) {
      const runner = race.runners[i];
      if (runner.scratched) {
        info(`Skipping scratched horse: ${runner.name}`);
        horseStats.push({ name: runner.name, error: 'Scratched' });
        continue;
      }
      info(`[${i + 1}/${race.runners.length}] Looking up: ${runner.name}`);
      const stats = await scrapeRacingZoneHorse(page, runner.name);
      horseStats.push(stats);
      if (i < race.runners.length - 1) await sleep(args.delayMs);
    }

    // Step 3: Save to Excel
    const outputFile = await saveToExcel(race, horseStats, outputPath);
    console.log(`\nDone! Excel file saved to: ${outputFile}`);

  } finally {
    await browser.close();
  }
}

main().catch((err) => {
  error('Fatal error:', err.message);
  process.exit(1);
});
