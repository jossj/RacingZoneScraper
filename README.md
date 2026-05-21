# Racing Zone Scraper

Scrapes a TAB race page for runner data, looks up each horse on RacingZone for
career statistics, and saves everything to an Excel workbook.

## Requirements

- Node.js 18+
- Google Chrome (Chromium) or Firefox installed
- Internet access

## Installation

```bash
npm install
npm run install-browsers
```

## Usage

```bash
node scraper.js
```

When you run the script it will prompt you to paste the TAB race URL:

```
============================================================
  Horse Racing Scraper
============================================================
Paste the full URL of the TAB race page and press Enter.
Example: https://www.tab.com.au/racing/meetings/RANDWICK/...

TAB race URL: <paste here>
```

### Options

| Option | Default | Description |
|---|---|---|
| `--output <path>` | `C:\tab\scrape` | Directory to save the Excel file |
| `--browser <name>` | `chromium` | `chromium` or `firefox` |
| `--no-headless` | — | Show the browser window (useful for debugging) |
| `--delay <seconds>` | `2` | Seconds to wait between RacingZone requests |

### Examples

```bash
# Run and enter the URL when prompted
node scraper.js

# Show browser window (good for debugging)
node scraper.js --no-headless

# Use Firefox and a custom output folder
node scraper.js --browser firefox --output "D:\racing"
```

## Output

An `.xlsx` file is created in the output folder containing four sheets:

| Sheet | Contents |
|---|---|
| **Race Info** | Venue, race name, distance, track condition, prize money, etc. |
| **Runners (TAB)** | All runners with barrier, jockey, trainer, weight, form, and odds |
| **Horse Stats (RacingZone)** | Career and L12M stats, breeding, stats by distance/condition/jockey/trainer |
| **Combined** | Runners sheet merged with stats sheet for a single view |

## Notes

- Playwright automatically manages the browser binary — no manual driver download needed.
- Run with `--no-headless` to watch the browser work or to debug selector issues.
- A 2-second polite delay is applied between each RacingZone request. Adjust with `--delay`.
- On Linux/Mac the Windows path `C:\tab\scrape` is automatically mapped to `~/tab/scrape`.
