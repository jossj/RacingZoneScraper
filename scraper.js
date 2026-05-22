'use strict';

const puppeteer = require('puppeteer');
const ExcelJS   = require('exceljs');
const readline  = require('readline');
const path      = require('path');
const fs        = require('fs');
const os        = require('os');

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
// Page helpers
// ---------------------------------------------------------------------------

async function scrollPage(page) {
  await page.evaluate(async () => {
    await new Promise(resolve => {
      let total = 0;
      const step = 800;
      const timer = setInterval(() => {
        window.scrollBy(0, step);
        total += step;
        if (total >= document.body.scrollHeight) {
          clearInterval(timer);
          window.scrollTo(0, 0);
          resolve();
        }
      }, 120);
    });
  });
  await sleep(1000);
}

async function waitForContent(page) {
  const selectors = [
    '[id^="runner-"]',
    '[class*="runner-card"]',
    '[class*="RunnerCard"]',
    '[class*="runner-row"]',
    '[class*="RunnerRow"]',
    '[class*="competitor"]',
    '[class*="Competitor"]',
    '[class*="form-runner"]',
  ];
  for (const sel of selectors) {
    try {
      await page.waitForSelector(sel, { timeout: 8000 });
      info(`Content detected via: ${sel}`);
      return sel;
    } catch { /* try next */ }
  }
  warn('Could not detect runner elements — scraping anyway after delay');
  await sleep(5000);
  return null;
}

async function dismissOverlays(page) {
  // Dismiss cookie banners, login modals, age verification, etc.
  const overlayBtns = [
    '[class*="cookie"] button[class*="accept"]',
    '[class*="cookie"] button[class*="close"]',
    '[class*="cookie"] button[class*="agree"]',
    '[id*="cookie"] button',
    '[class*="modal"] button[class*="close"]',
    '[aria-label*="close" i]',
    '[aria-label*="accept" i]',
    'button[class*="dismiss"]',
    'button[class*="consent"]',
  ];
  for (const sel of overlayBtns) {
    try {
      const btn = await page.$(sel);
      if (btn) { await btn.click(); await sleep(500); info(`Dismissed overlay: ${sel}`); }
    } catch { /* ignore */ }
  }
}

async function expandAllRunners(page) {
  // Click all "Show Form" / "Expand" toggles so full form details are visible
  const expanded = await page.evaluate(() => {
    let count = 0;
    const toggles = [
      ...document.querySelectorAll('[aria-expanded="false"]'),
      ...document.querySelectorAll('[class*="expand"]:not([class*="expanded"])'),
      ...document.querySelectorAll('[class*="show-form"]'),
      ...document.querySelectorAll('[class*="toggle"][class*="form"]'),
    ];
    // Deduplicate
    const seen = new Set();
    for (const el of toggles) {
      if (seen.has(el)) continue;
      seen.add(el);
      try { el.click(); count++; } catch { /* ignore */ }
    }
    return count;
  });
  if (expanded > 0) {
    info(`Clicked ${expanded} expand toggle(s)`);
    await sleep(2000);
  }
}

// ---------------------------------------------------------------------------
// Try to extract structured JSON data embedded by Next.js / Nuxt / etc.
// ---------------------------------------------------------------------------

async function tryExtractEmbeddedData(page) {
  return page.evaluate(() => {
    // Next.js
    if (window.__NEXT_DATA__) {
      try { return { source: 'next', data: window.__NEXT_DATA__ }; } catch { /* ignore */ }
    }
    // Nuxt
    if (window.__NUXT__) {
      try { return { source: 'nuxt', data: window.__NUXT__ }; } catch { /* ignore */ }
    }
    // Common SPA state keys
    for (const key of ['__INITIAL_STATE__', '__PRELOADED_STATE__', '__APP_STATE__', '__DATA__']) {
      if (window[key]) {
        try { return { source: key, data: window[key] }; } catch { /* ignore */ }
      }
    }
    // Look for <script id="__NEXT_DATA__"> tags
    const el = document.querySelector('#__NEXT_DATA__');
    if (el) {
      try { return { source: 'script-tag', data: JSON.parse(el.textContent) }; } catch { /* ignore */ }
    }
    return null;
  });
}

// ---------------------------------------------------------------------------
// Race info
// ---------------------------------------------------------------------------

async function getRaceInfo(page, url) {
  return page.evaluate((pageUrl) => {
    const info = { url: pageUrl, raceName: '', venue: '', date: '', raceNum: '', distance: '', raceClass: '', trackCondition: '' };

    const bodyText = document.body.innerText || '';

    // Venue - look for common Australian track names or header text
    const venueSelectors = [
      '[class*="venue"]', '[class*="Venue"]',
      '[class*="meeting-name"]', '[class*="MeetingName"]',
      '[class*="track-name"]', '[class*="TrackName"]',
      'h1', 'h2',
    ];
    for (const sel of venueSelectors) {
      const el = document.querySelector(sel);
      if (el) {
        const t = el.textContent.trim();
        if (t.length > 2 && t.length < 60) { info.raceName = t; break; }
      }
    }

    // Race number
    const raceNumMatch = bodyText.match(/\bRace\s*(\d+)\b/i);
    if (raceNumMatch) info.raceNum = raceNumMatch[1];

    // Distance
    const distMatch = bodyText.match(/\b(\d{3,4}m)\b/i);
    if (distMatch) info.distance = distMatch[1];

    // Date — multiple formats
    const datePatterns = [
      /\b(\d{1,2}[\s\-\/]\w{3,9}[\s\-\/]\d{4})\b/,
      /\b(\w{3,9}[\s\-\/]\d{1,2}[\s\-\/,]\s*\d{4})\b/,
      /\b(\d{4}-\d{2}-\d{2})\b/,
      /\b(\d{1,2}\/\d{1,2}\/\d{2,4})\b/,
    ];
    for (const pat of datePatterns) {
      const m = bodyText.match(pat);
      if (m) { info.date = m[1]; break; }
    }

    // Track condition
    const condMatch = bodyText.match(/\b(Firm\s*\d*|Good\s*\d*|Soft\s*\d*|Heavy\s*\d*|Synthetic|Wet\s*\d*)\b/i);
    if (condMatch) info.trackCondition = condMatch[1].trim();

    // Race class
    const classMatch = bodyText.match(/\b(G[123]|Group\s*[123]|Listed|BM\s*\d+|Benchmark\s*\d+|MDN|Maiden|Handicap|Open|3YO|4YO|2YO|3-4YO|WFA|CL\d+|NMW)\b/i);
    if (classMatch) info.raceClass = classMatch[1].trim();

    // Collect detail items from likely header elements
    const detailEls = document.querySelectorAll(
      '[class*="race-info"] *,' +
      '[class*="RaceInfo"] *,' +
      '[class*="race-detail"] *,' +
      '[class*="RaceDetail"] *,' +
      '[class*="race-header"] *,' +
      '[class*="RaceHeader"] *'
    );
    const details = [...detailEls]
      .map(e => e.textContent.trim())
      .filter(t => t.length > 1 && t.length < 80 && !t.includes('\n'))
      .slice(0, 15);
    info.details = details.join(' | ');

    return info;
  }, url);
}

// ---------------------------------------------------------------------------
// Runner scraping
// ---------------------------------------------------------------------------

async function scrapeRunners(page) {
  // First: expand any collapsed form sections
  await expandAllRunners(page);

  return page.evaluate(() => {
    const results = [];

    // ── Find runner container elements ────────────────────────────────────────
    const containerCandidates = [
      // id-based (matches #runner-N fragment pattern in URL)
      [...document.querySelectorAll('[id^="runner-"]')],
      // class-based
      [...document.querySelectorAll('[class*="runner-card"], [class*="RunnerCard"]')],
      [...document.querySelectorAll('[class*="runner-row"], [class*="RunnerRow"]')],
      [...document.querySelectorAll('[class*="runner-item"], [class*="RunnerItem"]')],
      [...document.querySelectorAll('[class*="competitor-row"], [class*="CompetitorRow"]')],
      [...document.querySelectorAll('[class*="form-runner"], [class*="FormRunner"]')],
      // fallback: any element containing both a horse name indicator and jockey/trainer
      [...document.querySelectorAll('[class*="runner"]')]
        .filter(el => el.textContent.length > 30 && !el.querySelector('[class*="runner"]')),
    ];

    let runnerEls = [];
    for (const cands of containerCandidates) {
      if (cands.length >= 2) { runnerEls = cands; break; }
    }

    if (!runnerEls.length) return results;

    // Helper: get first non-empty text from a list of selectors within a root
    const getText = (root, sels) => {
      for (const s of sels) {
        try {
          const el = root.querySelector(s);
          if (el) {
            const t = el.textContent.trim();
            if (t) return t;
          }
        } catch { /* bad selector, ignore */ }
      }
      return '';
    };

    // Helper: get all matching texts
    const getAll = (root, sel) => {
      try {
        return [...root.querySelectorAll(sel)].map(e => e.textContent.trim()).filter(Boolean);
      } catch { return []; }
    };

    for (const el of runnerEls) {
      const elText = el.textContent || '';
      if (elText.length < 10) continue;

      // ── Number ────────────────────────────────────────────────────────────
      const number = getText(el, [
        '[class*="runner-number"]', '[class*="RunnerNumber"]',
        '[class*="saddle-cloth"]', '[class*="SaddleCloth"]',
        '[class*="cloth-number"]', '[class*="ClothNumber"]',
        '[class*="competitor-number"]', '[class*="number"]',
      ]);

      // ── Horse name ────────────────────────────────────────────────────────
      const name = getText(el, [
        '[class*="horse-name"]', '[class*="HorseName"]',
        '[class*="runner-name"]', '[class*="RunnerName"]',
        '[class*="competitor-name"]', '[class*="CompetitorName"]',
        '[class*="animal-name"]',
        'h1', 'h2', 'h3',
      ]);

      if (!name || name.length < 2) continue;

      // ── Barrier ───────────────────────────────────────────────────────────
      const barrier = getText(el, [
        '[class*="barrier"]', '[class*="Barrier"]',
        '[class*="gate"]', '[class*="Gate"]',
        '[class*="draw"]', '[class*="Draw"]',
      ]);

      // ── Jockey ────────────────────────────────────────────────────────────
      const jockey = getText(el, [
        '[class*="jockey-name"]', '[class*="JockeyName"]',
        '[class*="jockey"]', '[class*="Jockey"]',
        '[class*="rider"]', '[class*="Rider"]',
        '[class*="driver"]', '[class*="Driver"]',
      ]);

      // ── Trainer ───────────────────────────────────────────────────────────
      const trainer = getText(el, [
        '[class*="trainer-name"]', '[class*="TrainerName"]',
        '[class*="trainer"]', '[class*="Trainer"]',
      ]);

      // ── Weight ────────────────────────────────────────────────────────────
      const weight = getText(el, [
        '[class*="weight"]', '[class*="Weight"]',
        '[class*="handicap"]', '[class*="Handicap"]',
        '[class*="kg"]',
      ]);

      // ── Form string ───────────────────────────────────────────────────────
      const form = getText(el, [
        '[class*="form-string"]', '[class*="FormString"]',
        '[class*="form-figures"]', '[class*="FormFigures"]',
        '[class*="last-starts"]', '[class*="LastStarts"]',
        '[class*="form-guide"]',  '[class*="FormGuide"]',
        '[class*="recent-form"]', '[class*="RecentForm"]',
        '[class*="form-numbers"]','[class*="FormNumbers"]',
      ]);

      // ── Odds ──────────────────────────────────────────────────────────────
      const oddsEls = getAll(el, [
        '[class*="price"]', '[class*="Price"]',
        '[class*="odds"]',  '[class*="Odds"]',
        'button[class*="bet"]',
      ].join(', '));

      let winOdds   = '';
      let placeOdds = '';
      // Odds are usually the first two price buttons/spans
      const oddsNums = oddsEls
        .map(t => t.replace(/[^\d.$]/g, '').trim())
        .filter(t => /^\$?[\d]+\.?\d*$/.test(t));
      if (oddsNums[0]) winOdds   = oddsNums[0];
      if (oddsNums[1]) placeOdds = oddsNums[1];

      // Also try more specific win/place selectors
      const winEl   = getText(el, ['[class*="win-price"]','[class*="WinPrice"]','[class*="win-odds"]','[class*="WinOdds"]','[class*="fixed-win"]']);
      const placeEl = getText(el, ['[class*="place-price"]','[class*="PlacePrice"]','[class*="place-odds"]','[class*="PlaceOdds"]','[class*="fixed-place"]']);
      if (winEl)   winOdds   = winEl;
      if (placeEl) placeOdds = placeEl;

      // ── Career stats ──────────────────────────────────────────────────────
      const careerText = getText(el, [
        '[class*="career-stats"]', '[class*="CareerStats"]',
        '[class*="career-record"]','[class*="CareerRecord"]',
        '[class*="career"]',       '[class*="Career"]',
        '[class*="stats"]',        '[class*="Stats"]',
      ]);

      // Parse career: "20: 5-4-3 $125,000" or "20 5 4 3" etc.
      let careerStarts = '', careerWins = '', careerSeconds = '', careerThirds = '', prizeMoney = '';
      if (careerText) {
        const prizeMatch = careerText.match(/\$[\d,]+/);
        if (prizeMatch) prizeMoney = prizeMatch[0];
        const numsMatch = careerText.match(/(\d+)\D+(\d+)\D+(\d+)\D+(\d+)/);
        if (numsMatch) {
          careerStarts  = numsMatch[1];
          careerWins    = numsMatch[2];
          careerSeconds = numsMatch[3];
          careerThirds  = numsMatch[4];
        }
      }

      // Win % and Place %
      const winPct   = (elText.match(/Win[s]?\s*[:%]\s*([\d.]+\s*%?)/i)   || [])[1] || '';
      const placePct = (elText.match(/Place[s]?\s*[:%]\s*([\d.]+\s*%?)/i) || [])[1] || '';

      // ── Condition stats ───────────────────────────────────────────────────
      // Collect label:value pairs from stats sections
      const condStats = {};
      const statRows = el.querySelectorAll('[class*="stat-row"],[class*="StatRow"],[class*="condition-row"],[class*="ConditionRow"]');
      for (const row of statRows) {
        const label = getText(row, ['[class*="label"]','[class*="Label"]','dt','th']);
        const value = getText(row, ['[class*="value"]','[class*="Value"]','dd','td']);
        if (label && value) condStats[label] = value;
      }

      // Also look for dt/dd pairs
      const dts = [...el.querySelectorAll('dt')];
      const dds = [...el.querySelectorAll('dd')];
      dts.forEach((dt, i) => {
        if (dds[i]) condStats[dt.textContent.trim()] = dds[i].textContent.trim();
      });

      // Extract common condition categories from text if structured data missing
      const conditionKeys = [
        ['Distance', /Distance[:\s]+([\d\w\s%\-\/]+)/i],
        ['Track',    /Track[:\s]+([\d\w\s%\-\/]+)/i],
        ['Firm',     /Firm[:\s]+([\d\w\s%\-\/]+)/i],
        ['Good',     /Good[:\s]+([\d\w\s%\-\/]+)/i],
        ['Soft',     /Soft[:\s]+([\d\w\s%\-\/]+)/i],
        ['Heavy',    /Heavy[:\s]+([\d\w\s%\-\/]+)/i],
        ['Barrier',  /Barrier[:\s]+([\d\w\s%\-\/]+)/i],
        ['1st Up',   /1st\s*Up[:\s]+([\d\w\s%\-\/]+)/i],
        ['2nd Up',   /2nd\s*Up[:\s]+([\d\w\s%\-\/]+)/i],
        ['3rd Up',   /3rd\s*Up[:\s]+([\d\w\s%\-\/]+)/i],
      ];
      for (const [key, pat] of conditionKeys) {
        if (!condStats[key]) {
          const m = elText.match(pat);
          if (m) condStats[key] = m[1].trim().slice(0, 30);
        }
      }

      // ── Race history ──────────────────────────────────────────────────────
      const raceHistory = [];
      const histSelectors = [
        '[class*="race-history"]',  '[class*="RaceHistory"]',
        '[class*="last-starts"]',   '[class*="LastStarts"]',
        '[class*="form-history"]',  '[class*="FormHistory"]',
        '[class*="run-history"]',   '[class*="RunHistory"]',
        '[class*="past-starts"]',   '[class*="PastStarts"]',
        '[class*="race-results"]',  '[class*="RaceResults"]',
      ];

      let histContainer = null;
      for (const sel of histSelectors) {
        histContainer = el.querySelector(sel);
        if (histContainer) break;
      }

      if (histContainer) {
        // Try table rows first
        const tableRows = histContainer.querySelectorAll('tr:not(:first-child)');
        if (tableRows.length) {
          for (const row of tableRows) {
            const cells = [...row.querySelectorAll('td,th')].map(c => c.textContent.trim());
            if (cells.length >= 3 && cells.some(c => /\d/.test(c))) {
              raceHistory.push(parseHistoryRow(cells));
            }
          }
        } else {
          // Try generic row elements
          const rowEls = histContainer.querySelectorAll('[class*="row"],[class*="Row"],[class*="item"],[class*="Item"]');
          for (const row of rowEls) {
            const cells = [...row.querySelectorAll('[class*="cell"],[class*="Cell"],[class*="col"],[class*="Col"],span,div')]
              .filter(c => !c.querySelector('[class*="cell"],[class*="col"],span,div'))
              .map(c => c.textContent.trim())
              .filter(Boolean);
            if (cells.length >= 3) {
              raceHistory.push(parseHistoryRow(cells));
            }
          }
        }
      }

      // ── Scratched ─────────────────────────────────────────────────────────
      const scratched = /\bscratched\b|\bSCR\b/.test(elText);

      // ── Age / Sex / Colour ────────────────────────────────────────────────
      const ageSexColour = getText(el, [
        '[class*="age-sex"]', '[class*="AgeSex"]',
        '[class*="horse-details"]','[class*="HorseDetails"]',
        '[class*="details"]',
      ]);

      // ── Sire / Dam ────────────────────────────────────────────────────────
      const sire = getText(el, ['[class*="sire"]','[class*="Sire"]']);
      const dam  = getText(el, ['[class*="dam"]','[class*="Dam"]']);

      results.push({
        number:       number.replace(/\D/g, '').slice(0, 2) || '',
        name:         name.replace(/\s+/g, ' ').trim(),
        barrier:      barrier.replace(/\D/g, '').slice(0, 2) || '',
        jockey:       jockey.replace(/\s+/g, ' ').trim(),
        trainer:      trainer.replace(/\s+/g, ' ').trim(),
        weight:       weight.replace(/\s+/g, ' ').trim(),
        form,
        winOdds,
        placeOdds,
        careerStarts,
        careerWins,
        careerSeconds,
        careerThirds,
        prizeMoney,
        winPct,
        placePct,
        ageSexColour: ageSexColour.replace(/\s+/g, ' ').trim(),
        sire:         sire.replace(/\s+/g, ' ').trim(),
        dam:          dam.replace(/\s+/g, ' ').trim(),
        condStats,
        raceHistory,
        scratched,
      });
    }

    return results;

    // Inner helper: convert an array of cell strings into a structured race entry
    function parseHistoryRow(cells) {
      const row = { raw: cells.join(' | ') };

      for (const cell of cells) {
        // Date
        if (!row.date && /\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}/.test(cell))
          row.date = cell;
        // Distance
        if (!row.distance && /^\d{3,4}m$/i.test(cell.trim()))
          row.distance = cell.trim();
        // Track condition
        if (!row.condition && /^(Firm|Good|Soft|Heavy|Synthetic|Wet)\d*$/i.test(cell.trim()))
          row.condition = cell.trim();
        // Race class
        if (!row.raceClass && /^(G[123]|Listed|BM\d+|MDN|Maiden|CL\d+|Open|Hcp|WFA|\d+YO)/i.test(cell.trim()))
          row.raceClass = cell.trim();
        // Position (e.g. "1st" "2nd" "1/8")
        if (!row.position && (/^\d{1,2}(st|nd|rd|th)$/i.test(cell.trim()) || /^\d{1,2}\/\d{1,2}$/.test(cell.trim())))
          row.position = cell.trim();
        // Margin
        if (!row.margin && /^(\d+(\.\d+)?L|SH|NK|HD|NS|NOS|NECK|HEAD)$/i.test(cell.trim()))
          row.margin = cell.trim();
        // Time
        if (!row.time && /^\d:\d{2}\.\d{1,2}$/.test(cell.trim()))
          row.time = cell.trim();
        // Weight
        if (!row.weight && /^\d{2}(\.\d)?$/.test(cell.trim()) && Number(cell) >= 48 && Number(cell) <= 65)
          row.weight = cell.trim();
        // Odds / price
        if (!row.odds && /^\$?[\d]+\.?\d{0,2}$/.test(cell.trim()) && !row.weight)
          row.odds = cell.trim().replace('$', '');
      }

      // Venue is often the one string with mixed case letters that isn't a name field
      for (const cell of cells) {
        if (!row.venue && /^[A-Z][a-z]+/.test(cell) && cell.length > 2 && cell.length < 25
            && !row.date && !/[%:]/.test(cell)) {
          row.venue = cell;
          break;
        }
      }

      return row;
    }
  });
}

// ---------------------------------------------------------------------------
// Try to enrich runner data from embedded JSON (Next.js / API response)
// ---------------------------------------------------------------------------

function parseEmbeddedRunners(embedded) {
  if (!embedded) return null;
  try {
    // Flatten nested data objects looking for runner arrays
    const str = JSON.stringify(embedded.data);
    // Look for arrays that have horse/runner objects
    const matches = str.match(/"(?:horse|runner|competitor)Name"\s*:\s*"([^"]+)"/g);
    if (matches && matches.length >= 2) {
      info(`Embedded data contains ${matches.length} horse/runner name references`);
    }
  } catch { /* ignore */ }
  return null; // DOM scraping is the primary method; embedded is a diagnostic hint
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

  const safeVenue = (raceInfo.raceName || 'Ladbrokes').replace(/[^a-zA-Z0-9 _-]/g, '_').replace(/\s+/g, '_');
  const dateStr   = (raceInfo.date || new Date().toISOString().slice(0, 10)).replace(/[\s\/\-]/g, '');
  const raceTag   = raceInfo.raceNum ? `_R${raceInfo.raceNum}` : '';
  const filename  = `Ladbrokes_${safeVenue}${raceTag}_${dateStr}.xlsx`;
  const filepath  = path.join(outputPath, filename);

  info('Saving Excel:', filepath);
  const wb = new ExcelJS.Workbook();

  // ── Sheet 1: Race Info ──────────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('Race Info');
    ws.columns = [{ width: 22 }, { width: 55 }];
    styleHeader(ws.addRow(['Field', 'Value']), 2);
    [
      ['Race Name / Venue',  raceInfo.raceName       || ''],
      ['Date',               raceInfo.date           || ''],
      ['Race Number',        raceInfo.raceNum        || ''],
      ['Distance',           raceInfo.distance       || ''],
      ['Class',              raceInfo.raceClass      || ''],
      ['Track Condition',    raceInfo.trackCondition || ''],
      ['Details',            raceInfo.details        || ''],
      ['URL',                raceInfo.url            || ''],
      ['Scraped At',         new Date().toISOString().replace('T', ' ').slice(0, 19)],
    ].forEach(([k, v], i) => styleData(ws.addRow([k, v]), 2, i % 2 === 1));
  }

  // ── Sheet 2: Runners ───────────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('Runners');
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    const cols = ['No.', 'Horse', 'Barrier', 'Jockey', 'Trainer', 'Weight', 'Form', 'Win Odds', 'Place Odds', 'Age/Sex/Colour', 'Sire', 'Dam', 'Scratched'];
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

  // ── Sheet 3: Career Stats ──────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('Career Stats');
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    const cols = ['No.', 'Horse', 'Starts', 'Wins', '2nds', '3rds', 'Prize Money', 'Win %', 'Place %'];
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

  // ── Sheet 4: Condition Stats ───────────────────────────────────────────
  {
    const ws = wb.addWorksheet('Condition Stats');
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    // Collect all condition stat keys used across all runners
    const allKeys = new Set();
    for (const r of runners) Object.keys(r.condStats || {}).forEach(k => allKeys.add(k));
    const keyList = [...allKeys];
    const cols = ['No.', 'Horse', ...keyList];
    styleHeader(ws.addRow(cols), cols.length);
    runners.forEach((r, i) => {
      styleData(ws.addRow([
        r.number, r.name,
        ...keyList.map(k => (r.condStats || {})[k] || ''),
      ]), cols.length, i % 2 === 1);
    });
    autoWidth(ws);
  }

  // ── Sheet 5: Race History ──────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('Race History');
    ws.views = [{ state: 'frozen', ySplit: 1, xSplit: 2 }];
    const cols = ['No.', 'Horse', 'Date', 'Venue', 'Distance', 'Condition', 'Class', 'Position', 'Margin', 'Time', 'Weight', 'Odds', 'Raw'];
    styleHeader(ws.addRow(cols), cols.length);

    let sheetRow = 2;
    for (const r of runners) {
      if (!r.raceHistory || !r.raceHistory.length) continue;

      const sec = ws.addRow([r.number, r.name, ...Array(cols.length - 2).fill('')]);
      styleSection(sec, cols.length);
      sheetRow++;

      let alt = false;
      for (const entry of r.raceHistory) {
        styleData(ws.addRow([
          r.number, r.name,
          entry.date      || '',
          entry.venue     || '',
          entry.distance  || '',
          entry.condition || '',
          entry.raceClass || '',
          entry.position  || '',
          entry.margin    || '',
          entry.time      || '',
          entry.weight    || '',
          entry.odds      || '',
          entry.raw       || '',
        ]), cols.length, alt);
        alt = !alt;
        sheetRow++;
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

    info('Dismissing overlays...');
    await dismissOverlays(page);

    info('Waiting for runner content...');
    await waitForContent(page);
    await sleep(2000);

    info('Scrolling page to trigger lazy loading...');
    await scrollPage(page);

    // Check for embedded JSON data (diagnostic)
    const embedded = await tryExtractEmbeddedData(page);
    if (embedded) {
      info(`Embedded data found (source: ${embedded.source})`);
      parseEmbeddedRunners(embedded);
    }

    info('Extracting race info...');
    const raceInfo = await getRaceInfo(page, url);
    info(`Race: ${raceInfo.raceName || 'Unknown'} | Date: ${raceInfo.date || 'N/A'} | Dist: ${raceInfo.distance || 'N/A'} | Cond: ${raceInfo.trackCondition || 'N/A'}`);

    info('Scraping runners...');
    const runners = await scrapeRunners(page);
    info(`Found ${runners.length} runner(s)`);

    if (!runners.length) {
      error('No runners found. The page may require login or use an unsupported layout.');
      error('Page title:', await page.title());
      info('Taking a diagnostic screenshot...');
      const diagPath = path.join(resolveOutputPath(args.output), 'diagnostic.png');
      fs.mkdirSync(path.dirname(diagPath), { recursive: true });
      await page.screenshot({ path: diagPath, fullPage: true });
      info('Screenshot saved to:', diagPath);
      process.exit(1);
    }

    const active    = runners.filter(r => !r.scratched);
    const scratched = runners.filter(r => r.scratched);
    info(`Runners: ${runners.length} total (${active.length} active, ${scratched.length} scratched)`);
    runners.forEach(r => info(`  #${r.number || '?'} ${r.name} | J: ${r.jockey || '–'} | T: ${r.trainer || '–'} | Win: ${r.winOdds || '–'} | Form: ${r.form || '–'}`));

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
