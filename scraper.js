'use strict';

const puppeteer          = require('puppeteer');
const ExcelJS            = require('exceljs');
const readline           = require('readline');
const path               = require('path');
const fs                 = require('fs');
const os                 = require('os');

const RACINGZONE_HORSES_URL = 'https://www.racingzone.com.au/horses/';
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

async function fillInput(page, selector, value) {
  // Works for plain HTML and React/Vue controlled inputs.
  // 1) Click to focus, 2) select-all + delete any existing text,
  // 3) set value via native setter (triggers React's synthetic onChange),
  // 4) dispatch real browser events so the framework picks up the change,
  // 5) type the text character-by-character as a fallback belt-and-suspenders.
  await page.click(selector);
  await page.keyboard.down('Control');
  await page.keyboard.press('KeyA');
  await page.keyboard.up('Control');
  await page.keyboard.press('Backspace');

  await page.evaluate((sel, val) => {
    const el = document.querySelector(sel);
    if (!el) return;
    // Native setter bypasses React's read-only descriptor
    const nativeSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    nativeSetter.call(el, val);
    el.dispatchEvent(new Event('input',  { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
  }, selector, value);
}

async function scrapeRacingZoneHorse(page, horseName) {
  const stats = {
    name: horseName, error: '',
    careerStarts: '', careerWins: '', careerSeconds: '', careerThirds: '',
    careerWinPct: '', careerPlacePct: '', careerPrizeMoney: '',
    l12mStarts: '', l12mWins: '', l12mSeconds: '', l12mThirds: '',
    statsByDistance: '', statsByCondition: '', statsByTrackType: '',
    statsByJockey: '', statsByTrainer: '',
    raceHistory: [],   // populated from OCR
  };

  info(`Searching RacingZone for: ${horseName}`);
  try {
    await page.goto(RACINGZONE_HORSES_URL, { waitUntil: 'networkidle2', timeout: 60000 });
    info(`  Landed on: ${page.url()}`);

    // Find the search input — first visible match wins
    const inputSels = [
      "input[name='horse_name']", "input[name='name']",
      "#horse_name", "#name",
      "input[placeholder*='horse' i]", "input[placeholder*='Find' i]",
      "input[placeholder*='Search' i]", "input[type='search']",
      "form input[type='text']", "input[type='text']",
    ];

    let inputSel = null;
    for (const sel of inputSels) {
      try {
        await page.waitForSelector(sel, { visible: true, timeout: 3000 });
        inputSel = sel;
        info(`  Search input found: ${sel}`);
        break;
      } catch { /* try next */ }
    }

    // Last resort: find any visible text input dynamically
    if (!inputSel) {
      inputSel = await page.evaluate(() => {
        const el = [...document.querySelectorAll('input')]
          .find(i => i.offsetParent !== null &&
                     (i.type === 'text' || i.type === 'search' || i.type === ''));
        if (!el) return null;
        if (el.id)   return `#${el.id}`;
        if (el.name) return `input[name="${el.name}"]`;
        return 'input[type="text"]';
      });
    }

    if (!inputSel) { stats.error = 'Search input not found'; return stats; }

    await fillInput(page, inputSel, horseName);
    await sleep(500);

    // Verify value landed
    const actual = await page.$eval(inputSel, el => el.value).catch(() => '');
    if (!actual) {
      warn(`  Input still empty — retrying with keyboard`);
      await page.click(inputSel, { clickCount: 3 });
      await page.keyboard.press('Backspace');
      await page.keyboard.type(horseName, { delay: 60 });
      await sleep(300);
    }
    info(`  Input value: "${await page.$eval(inputSel, el => el.value).catch(() => '?')}"`);

    // Submit — wait for either a full navigation or AJAX settling
    const btnSels = [
      "button[type='submit']", "input[type='submit']",
      "button[class*='search' i]", "form button",
    ];
    let submitBtn = null;
    for (const sel of btnSels) {
      submitBtn = await page.$(sel);
      if (submitBtn) break;
    }

    await Promise.all([
      page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }).catch(() => {}),
      submitBtn ? submitBtn.click() : page.keyboard.press('Enter'),
    ]);
    await sleep(1500);
    info(`  After search — URL: ${page.url()}`);

    // If we're on a results list, click the best-matching horse link
    const afterSearchUrl = page.url();
    const isOnHorsePage  = /\/horses?\/[^/]+\/?$/.test(afterSearchUrl);

    if (!isOnHorsePage) {
      // Gather all links visible on the page and find the best name match
      const allLinks = await page.evaluate((name) => {
        return [...document.querySelectorAll('a')]
          .filter(a => a.href && a.textContent.trim())
          .map(a => ({ text: a.textContent.trim(), href: a.href }));
      }, horseName);

      const nl = horseName.toLowerCase();
      const match = allLinks.find(l => l.text.toLowerCase().includes(nl))
                 || allLinks.find(l => nl.includes(l.text.toLowerCase()) && l.text.length > 3);

      if (match) {
        info(`  Clicking result: "${match.text}" → ${match.href}`);
        await Promise.all([
          page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }).catch(() => {}),
          page.goto(match.href, { waitUntil: 'networkidle2', timeout: 30000 }),
        ]);
        await sleep(1500);
      } else {
        // Log what was found so the user can see
        info(`  No matching link — links on page: ${allLinks.slice(0, 10).map(l => l.text).join(', ')}`);
      }
    }

    // Scroll the full page to force lazy-loaded content to render, then back to top.
    await page.evaluate(async () => {
      const step = 600;
      for (let y = 0; y < document.body.scrollHeight; y += step) {
        window.scrollTo(0, y);
        await new Promise(r => setTimeout(r, 80));
      }
      window.scrollTo(0, 0);
    });
    await sleep(1500);

    // Read visible page text directly from the DOM — no screenshot, no OCR,
    // no character errors. innerText returns text in visual reading order,
    // matching the structure the text parser expects.
    info(`  Extracting data from DOM: ${page.url()}`);
    const pageText = await page.evaluate(() => document.body.innerText || document.body.textContent || '');
    info(`  Page text: ${pageText.length} chars`);
    parsePageText(pageText, stats);

    // Overlay with structured table extraction for race history — direct DOM
    // column mapping is more accurate than regex parsing on running text.
    const domRows = await extractRaceHistoryDom(page);
    if (domRows.length > 0) {
      stats.raceHistory = domRows;
      info(`  DOM table: ${domRows.length} race history rows`);
    }

  } catch (err) {
    error(`RacingZone error for ${horseName}:`, err.message);
    stats.error = err.message.slice(0, 200);
  }
  return stats;
}

// ---------------------------------------------------------------------------
// DOM extraction helpers for RacingZone
// ---------------------------------------------------------------------------

function parseRaceRow(line) {
  // Accept dd/mm/yy, dd-mm-yy, dd.mm.yy  OR  "15 May 25" / "15 May 2025"
  const dateMatch =
    line.match(/\b(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})\b/) ||
    line.match(/\b(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s,]+\d{2,4})\b/i);
  if (!dateMatch) return null;

  const row = { date: dateMatch[1], raw: line };

  // Distance: 800m – 3200m
  const dist = line.match(/\b(\d{3,4}m)\b/i);
  if (dist) row.distance = dist[1];

  // Track condition: Good4, Soft7, Heavy10, Firm, Synthetic
  const cond = line.match(/\b(Firm\d*|Good\d*|Soft\d*|Heavy\d*|Synthetic|Syn)\b/i);
  if (cond) row.condition = cond[1];

  // Placing: "1st" / "2nd" / "3rd" / "1/8" (position of X runners)
  const placeSlash = line.match(/\b(\d{1,2})\/(\d{1,2})\b/);
  const placeSuffix = line.match(/\b(\d+(?:st|nd|rd|th))\b/i);
  if (placeSlash)       row.placing = `${placeSlash[1]}/${placeSlash[2]}`;
  else if (placeSuffix) row.placing = placeSuffix[1];

  // Race time: 1:09.50 or 0:57.30
  const time = line.match(/\b(\d:\d{2}\.\d{1,2})\b/);
  if (time) row.time = time[1];

  // Margin: 0.5L, 2.5L, SH, NK, LH, HD, NS
  const margin = line.match(/\b(\d+(?:\.\d+)?L|(?:SH|NK|LH|HD|NS|NOS|NECK|HEAD))\b/i);
  if (margin) row.margin = margin[1];

  // Race class: G1 G2 G3, Listed, BM64, MDN, CL1, 2YO, WFA, Open, Hcp
  const cls = line.match(/\b(G[123]|Gr[123]|Listed|BM\d+|MDN|CL\d+|\d+YO|WFA|Open|Hcp|HCP|FM\d*)\b/i);
  if (cls) row.class = cls[1];

  // Odds / SP: $3.50 or 3.50 near end of line
  const odds = line.match(/\$?(\d{1,3}\.\d{1,2})\s*$/);
  if (odds) row.odds = odds[1];

  // Weight carried: 54.0 kg, 57.5 kg — a decimal number between 48–65 not already captured as odds
  const weightMatch = line.match(/\b((?:4[89]|5\d|6[0-5])(?:\.\d)?)\b/g) || [];
  const notOdds = weightMatch.filter(w => row.odds !== w);
  if (notOdds.length) row.weight = notOdds[0];

  return row;
}

// Keywords that indicate a race history section on RacingZone (case-insensitive).
// Deliberately excludes nav-menu phrases like "Form Guide", "Recent Form", "Race Form".
const HISTORY_HEADINGS = [
  'race history', 'racing history', 'past runs', 'run history', 'recent runs',
  'form history', 'last starts', 'race record', 'past performances', 'run record',
  'recent starts', 'race results', 'last runs', 'past starts', 'last performances',
  'performance history', 'starts history', 'form record', 'runs',
];

// Label keywords whose value sits on the very next line in definition-list style.
const ADJACENT_LABELS = {
  sire: 'sire', dam: 'dam', colour: 'colour', color: 'colour',
  sex: 'sex', gender: 'sex', age: 'age', trainer: 'trainer',
  owner: 'owner', breeder: 'breeder', country: 'country', origin: 'country',
};

function parsePageText(rawText, stats) {
  const lines = rawText.split('\n').map(l => l.trim()).filter(Boolean);

  info(`  Page text preview: ${rawText.slice(0, 300).replace(/\n/g, ' ↵ ')}`);

  // Phase 1 — scan every line for profile fields, career stats, section headings
  let currentSection = null;
  let inHistory      = false;
  const sections = { distance: [], condition: [], jockey: [], trainer: [] };

  for (const line of lines) {
    const lower = line.toLowerCase();

    // ── Detect race history section start ────────────────────────────────────
    // Guard: line must be short (≤40 chars) so nav-bar sentences don't match.
    if (!inHistory && line.length <= 40 && HISTORY_HEADINGS.some(h => lower.includes(h))) {
      inHistory = true;
      currentSection = null;
      info(`  Race history section detected on line: "${line}"`);
      continue;
    }

    // ── Inside race history — collect rows then stop at next major section ───
    if (inHistory) {
      if (lower.match(/^(stats by|by distance|by condition|by track|by jockey|by trainer|career record)\b/)) {
        inHistory = false;
        // fall through to section-heading handling below
      } else {
        const row = parseRaceRow(line);
        if (row) stats.raceHistory.push(row);
        continue;
      }
    }

    // ── Key: Value pairs (Sire, Dam, Colour, etc.) ──────────────────────────
    // Also accept semicolons — OCR commonly misreads ':' as ';'.
    const kv = line.match(/^([A-Za-z][\w &/]+?)[;:]\s*(.+)$/);
    if (kv) {
      const key = kv[1].toLowerCase().trim();
      const val = kv[2].trim();
      if (key.includes('sire'))                         stats.sire    = stats.sire    || val;
      if (key.includes('dam'))                          stats.dam     = stats.dam     || val;
      if (key.includes('colour') || key === 'color')    stats.colour  = stats.colour  || val;
      if (key === 'sex' || key === 'gender')            stats.sex     = stats.sex     || val;
      if (key === 'age')                                stats.age     = stats.age     || val;
      if (key.includes('trainer'))                      stats.trainer = stats.trainer || val;
      if (key.includes('owner'))                        stats.owner   = stats.owner   || val;
      if (key.includes('breeder'))                      stats.breeder = stats.breeder || val;
      if (key.includes('country') || key === 'origin')  stats.country = stats.country || val;
    }

    // ── Career stats row ─────────────────────────────────────────────────────
    // Handles: "Career: 10 3 2 1", "Career 10-3-2-1", "Overall: 10: 3-2-1 $50,000"
    const isCareerLine = lower.includes('career') || lower.includes('overall record') || lower.includes('total starts');
    if (isCareerLine && !lower.includes('prize')) {
      // Try dash-separated record first: 10-3-2-1
      const dashRecord = line.match(/\b(\d+)-(\d+)-(\d+)-(\d+)\b/);
      const nums       = line.match(/\d[\d,]*/g) || [];
      const prize      = line.match(/\$[\d,]+/);
      if (dashRecord) {
        if (!stats.careerStarts)  stats.careerStarts  = dashRecord[1];
        if (!stats.careerWins)    stats.careerWins    = dashRecord[2];
        if (!stats.careerSeconds) stats.careerSeconds = dashRecord[3];
        if (!stats.careerThirds)  stats.careerThirds  = dashRecord[4];
      } else if (nums.length >= 4) {
        if (!stats.careerStarts)  stats.careerStarts  = nums[0];
        if (!stats.careerWins)    stats.careerWins    = nums[1];
        if (!stats.careerSeconds) stats.careerSeconds = nums[2];
        if (!stats.careerThirds)  stats.careerThirds  = nums[3];
      }
      if (prize && !stats.careerPrizeMoney) stats.careerPrizeMoney = prize[0];
    }

    // ── Last 12 months row ──────────────────────────────────────────────────
    const isL12mLine = lower.includes('last 12') || lower.includes('l12') ||
                       lower.includes('12 month') || lower.includes('past 12');
    if (isL12mLine && !stats.l12mStarts) {
      const dashRecord = line.match(/\b(\d+)-(\d+)-(\d+)-(\d+)\b/);
      const nums       = line.match(/\d[\d,]*/g) || [];
      if (dashRecord) {
        stats.l12mStarts  = dashRecord[1];
        stats.l12mWins    = dashRecord[2];
        stats.l12mSeconds = dashRecord[3];
        stats.l12mThirds  = dashRecord[4];
      } else if (nums.length >= 4) {
        stats.l12mStarts  = nums[0];
        stats.l12mWins    = nums[1];
        stats.l12mSeconds = nums[2];
        stats.l12mThirds  = nums[3];
      }
    }

    // ── Win % / Place % ─────────────────────────────────────────────────────
    const winPctMatch   = line.match(/win[s]?\s*[:%]\s*([\d.]+\s*%?)/i);
    const placePctMatch = line.match(/place[s]?\s*[:%]\s*([\d.]+\s*%?)/i);
    if (winPctMatch   && !stats.careerWinPct)   stats.careerWinPct   = winPctMatch[1].trim();
    if (placePctMatch && !stats.careerPlacePct) stats.careerPlacePct = placePctMatch[1].trim();

    // ── Stats-by-X section headings ──────────────────────────────────────────
    if      (lower.includes('by distance') || (lower.includes('distance') && lower.length < 25))  currentSection = 'distance';
    else if (lower.includes('by condition') || (lower.includes('condition') && lower.length < 25)) currentSection = 'condition';
    else if (lower.includes('by jockey')   || (lower.includes('jockey')   && lower.length < 25))  currentSection = 'jockey';
    else if (lower.includes('by trainer')  || (lower.includes('trainer')  && lower.length < 25))  currentSection = 'trainer';
    else if (currentSection) {
      if (line.length < 30 && /^[A-Z]/.test(line) && !/\d/.test(line)) {
        currentSection = null;
      } else {
        sections[currentSection].push(line);
      }
    }
  }

  if (sections.distance.length)  stats.statsByDistance  = sections.distance.join('\n');
  if (sections.condition.length) stats.statsByCondition = sections.condition.join('\n');
  if (sections.jockey.length)    stats.statsByJockey    = sections.jockey.join('\n');
  if (sections.trainer.length)   stats.statsByTrainer   = sections.trainer.join('\n');

  // Fallback: no section heading was detected — scan every line for anything
  // that looks like a race row (date + at least one other racing field).
  if (stats.raceHistory.length === 0) {
    info(`  No race history heading found — running fallback line scan`);
    for (const line of lines) {
      const row = parseRaceRow(line);
      if (row && (row.distance || row.condition || row.class || row.placing || row.time)) {
        stats.raceHistory.push(row);
      }
    }
    if (stats.raceHistory.length > 0)
      info(`  Fallback scan found ${stats.raceHistory.length} race rows`);
  }

  // Phase 2 — adjacent-line key/value (definition-list style: label on one
  // line, value on the next, no colon separator between them).
  for (let i = 0; i < lines.length - 1; i++) {
    const key = lines[i].toLowerCase().replace(/:$/, '').trim();
    const field = ADJACENT_LABELS[key];
    if (field && !stats[field]) stats[field] = lines[i + 1];
  }

  info(`  Parsed — career: ${stats.careerStarts}/${stats.careerWins}/${stats.careerSeconds}/${stats.careerThirds} | sire: ${stats.sire || '–'} | history rows: ${stats.raceHistory.length}`);
}

// ---------------------------------------------------------------------------
// Structured race-history extraction from the live DOM
// ---------------------------------------------------------------------------

async function extractRaceHistoryDom(page) {
  return page.evaluate(() => {
    const txt    = el => (el ? el.textContent.trim() : '');
    const low    = s  => s.toLowerCase();
    const dateRe = /\b\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}\b|\b\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i;

    // ── Strategy 1: HTML <table> whose cells contain race dates ──────────────
    const histTable = [...document.querySelectorAll('table')].find(t =>
      [...t.querySelectorAll('td')].some(td => dateRe.test(td.textContent))
    );

    if (histTable) {
      // Map <thead> columns to known field names
      const headers = [...histTable.querySelectorAll('thead th, thead td, tr:first-child th')]
        .map(h => low(txt(h)));

      const col = (...keys) => {
        const i = headers.findIndex(h => keys.some(k => h.includes(k)));
        return i !== -1 ? i : -1;
      };

      const cDate    = col('date');
      const cPlace   = col('pos', 'plac', 'finish', 'place');
      const cVenue   = col('venue', 'track', 'course');
      const cClass   = col('class', 'grade');
      const cDist    = col('dist', 'distance');
      const cWeight  = col('weight', 'wgt', 'kg');
      const cBarrier = col('barrier', 'gate');
      const cOdds    = col('odds', 'sp', 'price');
      const cWinner  = col('winner', '1st', 'first');
      const cMargin  = col('margin', 'len');
      const cTime    = col('time');
      const cInRun   = col('in run', 'inrun', 'sect');

      const rows = [];
      histTable.querySelectorAll('tbody tr, tr').forEach(tr => {
        const cells = [...tr.querySelectorAll('td')].map(txt);
        if (!cells.length) return;

        const dateIdx = cDate !== -1 ? cDate : cells.findIndex(c => dateRe.test(c));
        if (dateIdx === -1) return;

        rows.push({
          date:      cells[dateIdx]                  || '',
          placing:   cPlace   !== -1 ? cells[cPlace]   : '',
          venue:     cVenue   !== -1 ? cells[cVenue]   : '',
          class:     cClass   !== -1 ? cells[cClass]   : '',
          distance:  cDist    !== -1 ? cells[cDist]    : '',
          weight:    cWeight  !== -1 ? cells[cWeight]  : '',
          barrier:   cBarrier !== -1 ? cells[cBarrier] : '',
          odds:      cOdds    !== -1 ? cells[cOdds]    : '',
          winner2nd: cWinner  !== -1 ? cells[cWinner]  : '',
          margin:    cMargin  !== -1 ? cells[cMargin]  : '',
          time:      cTime    !== -1 ? cells[cTime]    : '',
          inRun:     cInRun   !== -1 ? cells[cInRun]   : '',
        });
      });
      if (rows.length) return rows;
    }

    // ── Strategy 2: div/React row layout ─────────────────────────────────────
    // Find a container whose heading suggests race history, then read child rows.
    const histKws = [
      'race history', 'racing history', 'past runs', 'run history', 'recent runs',
      'last starts', 'race record', 'past performances', 'form history', 'runs',
    ];
    const allHeadings = [...document.querySelectorAll(
      'h1,h2,h3,h4,h5,h6,[class*="heading"],[class*="section-title"],[class*="tab-title"]'
    )];
    let histContainer = null;
    for (const h of allHeadings) {
      if (histKws.some(kw => low(txt(h)).includes(kw))) {
        histContainer = h.parentElement;
        break;
      }
    }

    if (!histContainer) return [];

    const rows = [];
    const rowEls = histContainer.querySelectorAll('[class*="row"],[class*="item"],[class*="entry"],[class*="result"]');
    for (const row of rowEls) {
      // Leaf-node cells only (avoid picking up nested container text)
      const cells = [...row.querySelectorAll('span,div,td')]
        .filter(el => el.children.length === 0)
        .map(txt)
        .filter(Boolean);

      const dateIdx = cells.findIndex(c => dateRe.test(c));
      if (dateIdx === -1) continue;

      // Without a header row, assign columns positionally based on content type.
      const distIdx    = cells.findIndex(c => /\b\d{3,4}m\b/i.test(c));
      const condIdx    = cells.findIndex(c => /\b(Firm|Good|Soft|Heavy|Synthetic)/i.test(c));
      const placeIdx   = cells.findIndex(c => /^\d+\/?(\d+)?$|^\d+(st|nd|rd|th)$/i.test(c));
      const timeIdx    = cells.findIndex(c => /\d:\d{2}\.\d/.test(c));

      rows.push({
        date:      cells[dateIdx]                           || '',
        placing:   placeIdx  !== -1 ? cells[placeIdx]      : '',
        distance:  distIdx   !== -1 ? cells[distIdx]       : '',
        condition: condIdx   !== -1 ? cells[condIdx]       : '',
        time:      timeIdx   !== -1 ? cells[timeIdx]       : '',
        venue: '', class: '', weight: '', barrier: '', odds: '', winner2nd: '', margin: '', inRun: '',
      });
    }
    return rows;
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

  // ── Sheet 7: RZ Race History ───────────────────────────────────────────
  const ws7 = wb.addWorksheet('RZ Race History');
  ws7.views = [{ state: 'frozen', ySplit: 1, xSplit: 1 }];
  const rzHistCols = [
    'Horse', 'Date', 'Distance', 'Condition', 'Class',
    'Placing', 'Time', 'Margin', 'Weight', 'Odds', 'Raw Line',
  ];
  styleHeader(ws7.addRow(rzHistCols), rzHistCols.length);

  for (const s of horseStats) {
    if (!s.raceHistory || !s.raceHistory.length) continue;

    const secRow = ws7.addRow([s.name, ...Array(rzHistCols.length - 1).fill('')]);
    styleSection(secRow, rzHistCols.length);

    let alt = false;
    for (const entry of s.raceHistory) {
      const r = ws7.addRow([
        s.name,
        entry.date      || '',
        entry.distance  || '',
        entry.condition || '',
        entry.class     || '',
        entry.placing   || '',
        entry.time      || '',
        entry.margin    || '',
        entry.weight    || '',
        entry.odds      || '',
        entry.raw       || '',
      ]);
      styleData(r, rzHistCols.length, alt);
      alt = !alt;
    }
  }
  autoWidth(ws7);

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

    info('Starting RacingZone lookups (DOM scraping)...');
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
