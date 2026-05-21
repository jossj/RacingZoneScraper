# Racing Zone Scraper

Scrapes a TAB race page for runner data, looks up each horse on RacingZone for
career statistics, and saves everything to an Excel workbook.

## Requirements

- Python 3.9+
- Google Chrome **or** Firefox installed on your machine
- Internet access

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```
python scraper.py <TAB_RACE_URL> [options]
```

### Arguments

| Argument | Default | Description |
|---|---|---|
| `url` | *(required)* | Full URL of the TAB race page |
| `--output` | `C:\tab\scrape` | Directory to save the Excel file |
| `--browser` | `chrome` | `chrome` or `firefox` |
| `--headless` | on | Run browser invisibly (default) |
| `--no-headless` | — | Show the browser window (useful for debugging) |
| `--delay` | `2.0` | Seconds to wait between RacingZone requests |

### Examples

```bash
# Scrape a race and save to the default location
python scraper.py "https://www.tab.com.au/racing/meetings/RANDWICK/2024-11-02/races/1"

# Show browser window (good for debugging)
python scraper.py "https://www.tab.com.au/..." --no-headless

# Use Firefox and a custom output folder
python scraper.py "https://www.tab.com.au/..." --browser firefox --output "D:\racing"

# On Linux/Mac — output goes to ~/tab/scrape automatically
python scraper.py "https://www.tab.com.au/..."
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

- The scraper uses `webdriver-manager` to automatically download the correct
  ChromeDriver or GeckoDriver — no manual driver installation needed.
- If TAB changes its page layout the CSS selectors in `scrape_tab_page()` /
  `_scrape_tab_runners()` may need updating. Run with `--no-headless` to
  inspect the page visually.
- A 2-second polite delay is applied between each RacingZone request to avoid
  overloading the server. Adjust with `--delay`.
- On Linux/Mac the Windows path `C:\tab\scrape` is automatically mapped to
  `~/tab/scrape`.
