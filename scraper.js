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
// JSON API data hunting
// ---------------------------------------------------------------------------

// Recursively search an object/array for an array that looks like runners/competitors.
function findRunnersArray(obj, depth = 0) {
  if (depth > 8 || !obj || typeof obj !== 'object') return null;

  if (Array.isArray(obj) && obj.length >= 2) {
    const first = obj[0];
    if (first && typeof first === 'object' && !Array.isArray(first)) {
      const keys = Object.keys(first).map(k => k.toLowerCase());
      const looksLikeRunner =
        keys.some(k => ['name','horsename','runnername','horse','runner','competitor','animal'].some(n => k.includes(n))) ||
        keys.some(k => ['number','cloth','saddle','barrier','gate','draw'].some(n => k.includes(n))) ||
        keys.some(k => ['jockey','rider','driver','trainer'].some(n => k.includes(n)));
      if (looksLikeRunner) return obj;
    }
  }

  if (!Array.isArray(obj)) {
    for (const val of Object.values(obj)) {
      const found = findRunnersArray(val, depth + 1);
      if (found) return found;
    }
  }
  return null;
}

// Recursively find a race/meeting info object.
function findRaceInfo(obj, depth = 0) {
  if (depth > 6 || !obj || typeof obj !== 'object' || Array.isArray(obj)) return null;
  const keys = Object.keys(obj).map(k => k.toLowerCase());
  const looksLikeRace =
    keys.some(k => ['venue','track','course','racename','meetingname'].some(n => k.includes(n))) &&
    keys.some(k => ['runner','competitor','field','horses'].some(n => k.includes(n)));
  if (looksLikeRace) return obj;
  for (const val of Object.values(obj)) {
    const found = findRaceInfo(val, depth + 1);
    if (found) return found;
  }
  return null;
}

// Read a value from an object using multiple possible key names (case-insensitive).
function pick(obj, ...candidates) {
  if (!obj || typeof obj !== 'object') return '';
  for (const key of candidates) {
    // Exact key first
    if (obj[key] !== undefined && obj[key] !== null && obj[key] !== '') return String(obj[key]);
    // Case-insensitive
    const match = Object.keys(obj).find(k => k.toLowerCase() === key.toLowerCase());
    if (match !== undefined && obj[match] !== undefined && obj[match] !== null && obj[match] !== '')
      return String(obj[match]);
  }
  return '';
}

// Dive into nested object to find a named field (e.g. runner.jockey.name)
function deepPick(obj, ...candidates) {
  if (!obj || typeof obj !== 'object') return '';
  const direct = pick(obj, ...candidates);
  if (direct) return direct;
  // Try one level deep (e.g. obj.jockey.name)
  for (const key of Object.keys(obj)) {
    if (obj[key] && typeof obj[key] === 'object' && !Array.isArray(obj[key])) {
      const nested = pick(obj[key], ...candidates);
      if (nested) return nested;
    }
  }
  return '';
}

function mapApiRunner(r) {
  // Resolve nested objects (e.g. r.jockey might be { name: '...' })
  const jockeyObj  = r.jockey  || r.rider  || r.driver  || {};
  const trainerObj = r.trainer || r.coach  || {};
  const horseObj   = r.horse   || r.runner || r.competitor || r.animal || {};
  const priceObj   = r.prices  || r.price  || r.odds     || {};
  const formObj    = r.form    || r.formGuide || {};
  const careerObj  = r.career  || r.careerStats || r.stats || {};

  const name = pick(r, 'name','horseName','runnerName','competitorName','animalName') ||
               pick(horseObj, 'name','horseName') || '';

  const jockey = pick(r, 'jockeyName','riderName','driverName') ||
                 pick(jockeyObj, 'name','fullName','jockeyName') || '';

  const trainer = pick(r, 'trainerName') ||
                  pick(trainerObj, 'name','fullName','trainerName') || '';

  const number = pick(r, 'number','runnerNumber','clothNumber','saddle','position',
                        'competitorNumber','no','programNumber','tabNo') || '';

  const barrier = pick(r, 'barrier','barrierNumber','gate','draw','gateNumber') || '';

  const weight = pick(r, 'weight','handicapWeight','carryWeight','kg','handicap') || '';

  const form = pick(r, 'form','recentForm','formString','formFigures','lastStarts',
                      'formGuide','runnerForm','recentResults') ||
               pick(formObj, 'figures','string','recentForm','lastStarts') || '';

  const winOdds   = pick(r, 'fixedWin','winPrice','winOdds','win','priceWin','fixedWinPrice') ||
                    pick(priceObj, 'win','fixedWin','winPrice') || '';
  const placeOdds = pick(r, 'fixedPlace','placePrice','placeOdds','place','pricePlace','fixedPlacePrice') ||
                    pick(priceObj, 'place','fixedPlace','placePrice') || '';

  // Career stats
  const careerStarts  = pick(r, 'starts','careerStarts','totalStarts') ||
                        pick(careerObj, 'starts','total') || '';
  const careerWins    = pick(r, 'wins','careerWins')    || pick(careerObj, 'wins','first')  || '';
  const careerSeconds = pick(r, 'seconds','careerSeconds','places') || pick(careerObj, 'seconds','second') || '';
  const careerThirds  = pick(r, 'thirds','careerThirds') || pick(careerObj, 'thirds','third') || '';
  const prizeMoney    = pick(r, 'prizeMoney','earnings','totalPrizeMoney','prizeMoneyTotal') ||
                        pick(careerObj, 'prizeMoney','earnings') || '';
  const winPct        = pick(r, 'winPct','winPercentage','winPercent') ||
                        pick(careerObj, 'winPct','winPercentage') || '';
  const placePct      = pick(r, 'placePct','placePercentage','placePercent') ||
                        pick(careerObj, 'placePct','placePercentage') || '';

  // Horse attributes
  const sire = deepPick(r, 'sire','sireName','father') ||
               pick(horseObj, 'sire','sireName') || '';
  const dam  = deepPick(r, 'dam','damName','mother') ||
               pick(horseObj, 'dam','damName') || '';
  const age  = pick(r, 'age') || pick(horseObj, 'age') || '';
  const sex  = pick(r, 'sex','gender') || pick(horseObj, 'sex','gender') || '';
  const colour = pick(r, 'colour','color') || pick(horseObj, 'colour','color') || '';
  const ageSexColour = [age, sex, colour].filter(Boolean).join(' ');

  // Condition stats — look for stats objects inside the runner
  const condStats = {};
  const statFields = [
    ['Distance', ['byDistance','distanceStats','distance']],
    ['Track',    ['byTrack','trackStats','track']],
    ['Firm',     ['firm','firmStats']],
    ['Good',     ['good','goodStats']],
    ['Soft',     ['soft','softStats']],
    ['Heavy',    ['heavy','heavyStats']],
    ['Barrier',  ['byBarrier','barrierStats','barrier']],
    ['1st Up',   ['firstUp','1stUp']],
    ['2nd Up',   ['secondUp','2ndUp']],
    ['3rd Up',   ['thirdUp','3rdUp']],
  ];
  for (const [label, keys] of statFields) {
    const val = pick(r, ...keys);
    if (val) condStats[label] = val;
  }

  // Race / form history
  const raceHistory = [];
  const histSource = r.raceHistory || r.lastStarts || r.formHistory || r.runHistory ||
                     r.lastRuns || r.pastRaces || r.history || r.starts ||
                     (formObj && (formObj.history || formObj.lastStarts || formObj.runs)) || [];
  if (Array.isArray(histSource)) {
    for (const entry of histSource) {
      if (!entry || typeof entry !== 'object') continue;
      raceHistory.push({
        date:      pick(entry, 'date','raceDate','startDate','meetingDate') || '',
        venue:     pick(entry, 'venue','track','course','meetingName','trackName') || '',
        distance:  pick(entry, 'distance','dist') || '',
        condition: pick(entry, 'condition','trackCondition','going','trackRating') || '',
        raceClass: pick(entry, 'class','raceClass','grade','category') || '',
        position:  pick(entry, 'position','placing','place','finish','result','finishingPosition') || '',
        margin:    pick(entry, 'margin','margins') || '',
        time:      pick(entry, 'time','raceTime','finishTime') || '',
        weight:    pick(entry, 'weight','handicap','carryWeight') || '',
        odds:      pick(entry, 'odds','sp','startingPrice','price') || '',
        jockey:    deepPick(entry, 'jockey','rider','jockeyName','riderName') || '',
        barrier:   pick(entry, 'barrier','barrierNumber','gate') || '',
        raw:       '',
      });
    }
  }

  const scratchedRaw = pick(r, 'scratched','isScratched','status','runnerStatus') || '';
  const scratched = /true|scr|scratched/i.test(scratchedRaw);

  return {
    number, name, barrier, jockey, trainer, weight, form,
    winOdds, placeOdds,
    careerStarts, careerWins, careerSeconds, careerThirds, prizeMoney, winPct, placePct,
    ageSexColour, sire, dam,
    condStats, raceHistory, scratched,
  };
}

function extractRaceInfoFromApi(apiObj) {
  if (!apiObj) return {};
  const info = {};
  const raceObj = findRaceInfo(apiObj) || apiObj;

  info.raceName      = pick(raceObj, 'raceName','name','meetingName','trackName','venue','course') || '';
  info.venue         = pick(raceObj, 'venue','track','course','venueName','meetingVenue') || info.raceName;
  info.date          = pick(raceObj, 'date','raceDate','meetingDate','startDate') || '';
  info.raceNum       = pick(raceObj, 'raceNumber','number','race','raceNo','num') || '';
  info.distance      = pick(raceObj, 'distance','dist') || '';
  info.raceClass     = pick(raceObj, 'class','raceClass','grade','category','raceGrade') || '';
  info.trackCondition = pick(raceObj, 'trackCondition','condition','going','trackRating','surface') || '';
  return info;
}

// ---------------------------------------------------------------------------
// DOM-based scraping (fallback when API intercept has no runner data)
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

async function getRaceInfoFromDom(page, url) {
  return page.evaluate((pageUrl) => {
    const bodyText = document.body.innerText || '';
    const info = { url: pageUrl, raceName: '', venue: '', date: '', raceNum: '', distance: '', raceClass: '', trackCondition: '' };

    // Race name / title
    for (const sel of ['h1','h2','[class*="race-name"]','[class*="RaceName"]','[class*="race-title"]']) {
      const el = document.querySelector(sel);
      if (el) { const t = el.textContent.trim(); if (t.length > 2 && t.length < 80) { info.raceName = t; break; } }
    }

    // Race number
    const raceNumM = bodyText.match(/\bRace\s*(\d+)\b/i);
    if (raceNumM) info.raceNum = raceNumM[1];

    // Distance
    const distM = bodyText.match(/\b(\d{3,4}m)\b/i);
    if (distM) info.distance = distM[1];

    // Date
    for (const pat of [
      /\b(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{4})\b/i,
      /\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})\b/,
      /\b(\d{4}-\d{2}-\d{2})\b/,
    ]) {
      const m = bodyText.match(pat);
      if (m) { info.date = m[1]; break; }
    }

    // Track condition
    const condM = bodyText.match(/\b(Firm\s*\d*|Good\s*\d*|Soft\s*\d*|Heavy\s*\d*|Synthetic|Wet\s*\d*)\b/i);
    if (condM) info.trackCondition = condM[1].trim();

    // Race class
    const clsM = bodyText.match(/\b(G[123]|Group\s*[123]|Listed|BM\s*\d+|Benchmark\s*\d+|MDN|Maiden|Handicap|Open|WFA|CL\d+)\b/i);
    if (clsM) info.raceClass = clsM[1].trim();

    return info;
  }, url);
}

async function scrapeRunnersFromDom(page) {
  return page.evaluate(() => {
    // ── 1. Locate runner container elements ───────────────────────────────
    const strategies = [
      // id-based (#runner-N fragments)
      () => [...document.querySelectorAll('[id^="runner-"]')],
      () => [...document.querySelectorAll('[id^="Runner-"]')],
      () => [...document.querySelectorAll('[id^="competitor-"]')],
      // class contains 'runner' and has siblings (repeated)
      () => {
        const candidates = [...document.querySelectorAll('[class*="runner"]')]
          .filter(el => !el.querySelector('[class*="runner"]') && el.textContent.length > 40);
        // Only accept if we found a reasonable count
        return candidates.length >= 2 ? candidates : [];
      },
      () => {
        const candidates = [...document.querySelectorAll('[class*="Runner"]')]
          .filter(el => !el.querySelector('[class*="Runner"]') && el.textContent.length > 40);
        return candidates.length >= 2 ? candidates : [];
      },
      () => [...document.querySelectorAll('[class*="competitor"]')]
            .filter(el => !el.querySelector('[class*="competitor"]') && el.textContent.length > 30),
      () => [...document.querySelectorAll('[class*="Competitor"]')]
            .filter(el => !el.querySelector('[class*="Competitor"]') && el.textContent.length > 30),
      // data-attribute based
      () => [...document.querySelectorAll('[data-runner-id],[data-runner],[data-horse-id],[data-competitor]')],
      // Table rows containing racing data
      () => [...document.querySelectorAll('tbody tr')]
            .filter(tr => tr.textContent.length > 30 && tr.cells.length >= 4),
    ];

    let runnerEls = [];
    for (const strategy of strategies) {
      try {
        const els = strategy();
        if (els.length >= 2) { runnerEls = els; break; }
      } catch { /* ignore */ }
    }

    if (!runnerEls.length) return [];

    // ── 2. Helper: get first non-empty text from selectors ────────────────
    const getText = (root, sels) => {
      for (const s of sels) {
        try {
          const el = root.querySelector(s);
          if (el) { const t = el.textContent.trim(); if (t) return t; }
        } catch { /* ignore */ }
      }
      return '';
    };

    const getAllTexts = (root, sel) => {
      try { return [...root.querySelectorAll(sel)].map(e => e.textContent.trim()).filter(Boolean); }
      catch { return []; }
    };

    // ── 3. Per-runner extraction ──────────────────────────────────────────
    const results = [];

    for (const el of runnerEls) {
      const elText = el.textContent || '';
      if (elText.trim().length < 5) continue;

      // Number
      const number = getText(el, [
        '[class*="number"i]','[class*="cloth"i]','[class*="saddle"i]',
        '[class*="RunnerNumber"i]','[class*="SaddleCloth"i]',
        '[data-number]','[data-cloth]',
      ]) || (el.id ? el.id.replace(/\D/g, '') : '');

      // Horse name — try many class patterns, then fall back to largest text in header
      let name = getText(el, [
        '[class*="horse-name"i]','[class*="HorseName"]',
        '[class*="runner-name"i]','[class*="RunnerName"]',
        '[class*="competitor-name"i]','[class*="CompetitorName"]',
        '[class*="animal-name"i]','[class*="AnimalName"]',
        '[data-horse-name]','[data-name]',
      ]);
      if (!name) {
        // Try headings
        for (const tag of ['h1','h2','h3','h4']) {
          const h = el.querySelector(tag);
          if (h) { name = h.textContent.trim(); break; }
        }
      }
      if (!name || name.length < 2) continue;

      // Barrier
      const barrier = getText(el, [
        '[class*="barrier"i]','[class*="Barrier"]',
        '[class*="gate"i]','[class*="Gate"]',
        '[class*="draw"i]','[class*="Draw"]',
        '[data-barrier]','[data-gate]',
      ]);

      // Jockey
      const jockey = getText(el, [
        '[class*="jockey-name"i]','[class*="JockeyName"]',
        '[class*="jockey"i]','[class*="Jockey"]',
        '[class*="rider"i]','[class*="Rider"]',
        '[class*="driver"i]','[class*="Driver"]',
        '[data-jockey]',
      ]);

      // Trainer
      const trainer = getText(el, [
        '[class*="trainer-name"i]','[class*="TrainerName"]',
        '[class*="trainer"i]','[class*="Trainer"]',
        '[data-trainer]',
      ]);

      // Weight
      const weight = getText(el, [
        '[class*="weight"i]','[class*="Weight"]',
        '[class*="handicap"i]','[class*="Handicap"]',
        '[data-weight]','[data-handicap]',
      ]);

      // Form string
      const form = getText(el, [
        '[class*="form-string"i]','[class*="FormString"]',
        '[class*="form-figures"i]','[class*="FormFigures"]',
        '[class*="recent-form"i]','[class*="RecentForm"]',
        '[class*="last-starts"i]','[class*="LastStarts"]',
        '[class*="form-numbers"i]','[class*="FormNumbers"]',
        '[data-form]','[data-form-string]',
      ]);

      // Odds — collect all price-like elements
      const priceEls = getAllTexts(el, [
        '[class*="price"i]','[class*="Price"]',
        '[class*="odds"i]','[class*="Odds"]',
        '[class*="win-price"i]','[class*="WinPrice"]',
        '[class*="fixed-win"i]','[class*="FixedWin"]',
        '[data-price]','[data-odds]',
      ].join(', '));

      let winOdds = '', placeOdds = '';
      const winEl = getText(el, [
        '[class*="win-price"i]','[class*="WinPrice"]',
        '[class*="win-odds"i]','[class*="WinOdds"]',
        '[class*="fixed-win"i]','[class*="FixedWin"]',
      ]);
      const placeEl = getText(el, [
        '[class*="place-price"i]','[class*="PlacePrice"]',
        '[class*="place-odds"i]','[class*="PlaceOdds"]',
        '[class*="fixed-place"i]','[class*="FixedPlace"]',
      ]);
      if (winEl)   winOdds   = winEl;
      if (placeEl) placeOdds = placeEl;
      if (!winOdds && priceEls[0]) winOdds   = priceEls[0];
      if (!placeOdds && priceEls[1]) placeOdds = priceEls[1];

      // Career stats text block
      const careerText = getText(el, [
        '[class*="career"i]','[class*="Career"]',
        '[class*="career-stats"i]','[class*="CareerStats"]',
        '[class*="career-record"i]','[class*="CareerRecord"]',
        '[class*="career-summary"i]',
      ]);
      let careerStarts = '', careerWins = '', careerSeconds = '', careerThirds = '', prizeMoney = '';
      if (careerText) {
        const prizeM = careerText.match(/\$[\d,]+/);
        if (prizeM) prizeMoney = prizeM[0];
        const numsM = careerText.match(/(\d+)\D+(\d+)\D+(\d+)\D+(\d+)/);
        if (numsM) {
          careerStarts  = numsM[1]; careerWins = numsM[2];
          careerSeconds = numsM[3]; careerThirds = numsM[4];
        }
      }
      const winPct   = (elText.match(/Win[s]?\s*[:%]\s*([\d.]+\s*%?)/i)   || [])[1] || '';
      const placePct = (elText.match(/Place[s]?\s*[:%]\s*([\d.]+\s*%?)/i) || [])[1] || '';

      // Sire / Dam / Age / Sex / Colour
      const sire = getText(el, ['[class*="sire"i]','[class*="Sire"]','[data-sire]']);
      const dam  = getText(el, ['[class*="dam"i]','[class*="Dam"]','[data-dam]']);
      const ageSexColour = getText(el, [
        '[class*="age-sex"i]','[class*="AgeSex"]',
        '[class*="horse-details"i]','[class*="HorseDetails"]',
        '[class*="pedigree"i]',
      ]);

      // Condition stats — dt/dd pairs or stat rows
      const condStats = {};
      const dtEls = [...el.querySelectorAll('dt')];
      const ddEls = [...el.querySelectorAll('dd')];
      dtEls.forEach((dt, i) => { if (ddEls[i]) condStats[dt.textContent.trim()] = ddEls[i].textContent.trim(); });

      const statRows = el.querySelectorAll('[class*="stat-row"i],[class*="StatRow"],[class*="condition-row"i]');
      for (const row of statRows) {
        const label = getText(row, ['[class*="label"i]','[class*="Label"]','th','dt']);
        const value = getText(row, ['[class*="value"i]','[class*="Value"]','td','dd']);
        if (label && value) condStats[label] = value;
      }

      // Race history
      const raceHistory = [];
      let histEl = null;
      for (const sel of [
        '[class*="race-history"i]','[class*="RaceHistory"]',
        '[class*="last-starts"i]','[class*="LastStarts"]',
        '[class*="form-history"i]','[class*="FormHistory"]',
        '[class*="run-history"i]','[class*="RunHistory"]',
        '[class*="past-starts"i]','[class*="form-guide"i]',
      ]) {
        histEl = el.querySelector(sel);
        if (histEl) break;
      }

      if (histEl) {
        // Try table rows
        const trs = histEl.querySelectorAll('tr:not(:first-child)');
        if (trs.length) {
          for (const tr of trs) {
            const cells = [...tr.querySelectorAll('td,th')].map(c => c.textContent.trim()).filter(Boolean);
            if (cells.length >= 3) raceHistory.push(parseHistoryRow(cells));
          }
        } else {
          // Generic row-like children
          const rowEls = histEl.querySelectorAll('[class*="row"i],[class*="item"i],[class*="entry"i]');
          for (const row of rowEls) {
            const leafNodes = [...row.querySelectorAll('span,td,li,p')]
              .filter(n => !n.querySelector('span,td,li,p'))
              .map(n => n.textContent.trim()).filter(Boolean);
            if (leafNodes.length >= 3) raceHistory.push(parseHistoryRow(leafNodes));
          }
        }
      }

      const scratched = /\bscratched\b|\bSCR\b/.test(elText);

      results.push({
        number:  (number || '').replace(/\D/g, '').slice(0, 2),
        name:    name.replace(/\s+/g, ' ').trim(),
        barrier: (barrier || '').replace(/\D/g, '').slice(0, 2),
        jockey:  jockey.replace(/\s+/g, ' ').trim(),
        trainer: trainer.replace(/\s+/g, ' ').trim(),
        weight:  weight.replace(/\s+/g, ' ').trim(),
        form, winOdds, placeOdds,
        careerStarts, careerWins, careerSeconds, careerThirds, prizeMoney, winPct, placePct,
        ageSexColour: ageSexColour.replace(/\s+/g, ' ').trim(),
        sire: sire.replace(/\s+/g, ' ').trim(),
        dam:  dam.replace(/\s+/g, ' ').trim(),
        condStats, raceHistory, scratched,
      });
    }

    return results;

    function parseHistoryRow(cells) {
      const row = { raw: cells.join(' | ') };
      for (const cell of cells) {
        const c = cell.trim();
        if (!row.date     && /\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}/.test(c)) { row.date = c; continue; }
        if (!row.distance && /^\d{3,4}m$/i.test(c))                            { row.distance = c; continue; }
        if (!row.condition && /^(Firm|Good|Soft|Heavy|Synthetic|Wet)\d*$/i.test(c)) { row.condition = c; continue; }
        if (!row.raceClass && /^(G[123]|Listed|BM\d+|MDN|CL\d+|Open|Hcp|WFA|\d+YO)/i.test(c)) { row.raceClass = c; continue; }
        if (!row.position  && (/^\d{1,2}(st|nd|rd|th)$/i.test(c) || /^\d{1,2}\/\d{1,2}$/.test(c))) { row.position = c; continue; }
        if (!row.margin    && /^(\d+(\.\d+)?L|SH|NK|HD|NS|NOS|NECK|HEAD)$/i.test(c)) { row.margin = c; continue; }
        if (!row.time      && /^\d:\d{2}\.\d{1,2}$/.test(c))                    { row.time = c; continue; }
        if (!row.weight    && /^\d{2}(\.\d)?$/.test(c) && +c >= 48 && +c <= 65) { row.weight = c; continue; }
        if (!row.odds      && /^\$?[\d]+\.?\d{0,2}$/.test(c))                   { row.odds = c.replace('$',''); }
      }
      // Venue: first title-case word-group not yet captured
      for (const cell of cells) {
        if (!row.venue && /^[A-Z][a-z]+(\s[A-Z][a-z]+)?$/.test(cell.trim()) && cell.length < 25) {
          row.venue = cell.trim(); break;
        }
      }
      return row;
    }
  });
}

// ---------------------------------------------------------------------------
// Diagnostic dump on failure
// ---------------------------------------------------------------------------

async function saveDiagnostics(page, outputPath, capturedJson) {
  fs.mkdirSync(outputPath, { recursive: true });

  // Screenshot
  const ssPath = path.join(outputPath, 'diagnostic.png');
  await page.screenshot({ path: ssPath, fullPage: true });
  info('Diagnostic screenshot:', ssPath);

  // HTML snippet (first 8000 chars of body)
  const bodyHtml = await page.evaluate(() => document.body.innerHTML.slice(0, 8000));
  const htmlPath = path.join(outputPath, 'diagnostic.html');
  fs.writeFileSync(htmlPath, bodyHtml);
  info('Diagnostic HTML:', htmlPath);

  // Captured API JSON
  if (capturedJson.length) {
    const jsonPath = path.join(outputPath, 'captured_api.json');
    fs.writeFileSync(jsonPath, JSON.stringify(capturedJson, null, 2));
    info('Captured API responses:', jsonPath);
  }
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
    const cols = ['No.','Horse','Date','Venue','Distance','Condition','Class','Position','Margin','Time','Weight','Odds','Jockey','Barrier','Raw'];
    styleHeader(ws.addRow(cols), cols.length);
    for (const r of runners) {
      if (!r.raceHistory || !r.raceHistory.length) continue;
      styleSection(ws.addRow([r.number, r.name, ...Array(cols.length - 2).fill('')]), cols.length);
      let alt = false;
      for (const e of r.raceHistory) {
        styleData(ws.addRow([
          r.number, r.name,
          e.date || '', e.venue || '', e.distance || '', e.condition || '',
          e.raceClass || '', e.position || '', e.margin || '', e.time || '',
          e.weight || '', e.odds || '', e.jockey || '', e.barrier || '', e.raw || '',
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

    // ── Intercept ALL JSON API responses before navigation ────────────────
    const capturedJson = [];
    page.on('response', async (response) => {
      try {
        const ct = response.headers()['content-type'] || '';
        if (response.status() === 200 && ct.includes('json')) {
          const json = await response.json();
          capturedJson.push({ url: response.url(), data: json });
        }
      } catch { /* binary or non-JSON — ignore */ }
    });

    info('Navigating to:', url);
    await page.goto(url, { waitUntil: 'networkidle2', timeout: 60000 });
    await sleep(2000);

    await dismissOverlays(page);
    await scrollPage(page);
    await expandAllRunners(page);
    await sleep(1000);

    // ── Attempt 1: parse runners from intercepted API data ────────────────
    let runners  = [];
    let raceInfo = { url };

    info(`Captured ${capturedJson.length} JSON API response(s) — searching for runner data...`);

    for (const resp of capturedJson) {
      const arr = findRunnersArray(resp.data);
      if (arr && arr.length >= 2) {
        info(`Runner data found in API response: ${resp.url} (${arr.length} runners)`);
        runners  = arr.map(mapApiRunner).filter(r => r.name.length > 0);
        const ri = extractRaceInfoFromApi(resp.data);
        raceInfo = Object.assign({ url }, ri);
        break;
      }
    }

    // ── Attempt 2: DOM scraping ───────────────────────────────────────────
    if (!runners.length) {
      warn('No runner data found in API responses — falling back to DOM scraping');
      runners = await scrapeRunnersFromDom(page);
    }

    // ── Fill in race info from DOM if still missing ───────────────────────
    const domRaceInfo = await getRaceInfoFromDom(page, url);
    for (const key of Object.keys(domRaceInfo)) {
      if (!raceInfo[key]) raceInfo[key] = domRaceInfo[key];
    }

    // ── Bail out with diagnostics if still empty ──────────────────────────
    if (!runners.length) {
      error('No runners found via API intercept or DOM scraping.');
      error('Page title:', await page.title());
      error('Saving diagnostic files...');
      await saveDiagnostics(page, outputPath, capturedJson);
      info('Check diagnostic.png and diagnostic.html in the output folder to inspect the page structure.');
      if (capturedJson.length) {
        info('Also check captured_api.json — the runner data may be there under an unexpected key.');
        info('API URLs captured:');
        capturedJson.forEach(r => info(' ', r.url));
      }
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
