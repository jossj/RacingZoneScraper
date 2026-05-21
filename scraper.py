"""
Horse Racing Scraper
--------------------
Scrapes race data from a TAB race page, then fetches horse statistics from
RacingZone for each runner, and saves the combined data to an Excel workbook.

Usage:
    python scraper.py <TAB_RACE_URL> [--output <path>] [--headless] [--browser chrome|firefox]

Example:
    python scraper.py "https://www.tab.com.au/racing/meetings/..." --headless
"""

import argparse
import logging
import os
import platform
import sys
import time
from dataclasses import dataclass, field, fields
from pathlib import Path
from typing import Optional

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

RACINGZONE_HORSES_URL = "https://www.racingzone.com.au/statistics/horses/"
DEFAULT_WAIT = 15  # seconds


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass
class Runner:
    number: str = ""
    name: str = ""
    barrier: str = ""
    jockey: str = ""
    trainer: str = ""
    weight: str = ""
    age: str = ""
    sex: str = ""
    form: str = ""
    win_odds: str = ""
    place_odds: str = ""
    scratched: bool = False


@dataclass
class RaceInfo:
    venue: str = ""
    race_number: str = ""
    race_name: str = ""
    date: str = ""
    time: str = ""
    distance: str = ""
    track_condition: str = ""
    prize_money: str = ""
    race_class: str = ""
    runners: list = field(default_factory=list)


@dataclass
class HorseStats:
    name: str = ""
    # Career totals
    career_starts: str = ""
    career_wins: str = ""
    career_seconds: str = ""
    career_thirds: str = ""
    career_win_pct: str = ""
    career_place_pct: str = ""
    career_prize_money: str = ""
    # Last 12 months
    l12m_starts: str = ""
    l12m_wins: str = ""
    l12m_seconds: str = ""
    l12m_thirds: str = ""
    # By distance (raw text)
    stats_by_distance: str = ""
    # By track condition (raw text)
    stats_by_condition: str = ""
    # By track type (raw text)
    stats_by_track_type: str = ""
    # Jockey stats (raw text)
    stats_by_jockey: str = ""
    # Trainer stats (raw text)
    stats_by_trainer: str = ""
    # Additional info
    sire: str = ""
    dam: str = ""
    colour: str = ""
    sex: str = ""
    age: str = ""
    country: str = ""
    owner: str = ""
    breeder: str = ""
    error: str = ""


# ---------------------------------------------------------------------------
# WebDriver factory
# ---------------------------------------------------------------------------

def build_driver(browser: str = "chrome", headless: bool = True) -> webdriver.Remote:
    browser = browser.lower()
    if browser == "firefox":
        opts = FirefoxOptions()
        if headless:
            opts.add_argument("--headless")
        opts.set_preference("dom.webnotifications.enabled", False)
        service = FirefoxService(GeckoDriverManager().install())
        return webdriver.Firefox(service=service, options=opts)

    # Default: Chrome
    opts = ChromeOptions()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=opts)
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"},
    )
    return driver


def wait_for(driver, by, selector, timeout=DEFAULT_WAIT):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, selector))
    )


def safe_text(el) -> str:
    try:
        return el.text.strip()
    except (StaleElementReferenceException, AttributeError):
        return ""


def safe_find(driver_or_el, by, selector) -> Optional[object]:
    try:
        return driver_or_el.find_element(by, selector)
    except NoSuchElementException:
        return None


def safe_find_all(driver_or_el, by, selector) -> list:
    try:
        return driver_or_el.find_elements(by, selector)
    except NoSuchElementException:
        return []


# ---------------------------------------------------------------------------
# TAB scraper
# ---------------------------------------------------------------------------

def scrape_tab_page(driver: webdriver.Remote, url: str) -> RaceInfo:
    log.info("Navigating to TAB page: %s", url)
    driver.get(url)
    time.sleep(3)

    race = RaceInfo()

    # --- Race header / metadata ---
    try:
        # Wait for page to stabilise - look for a common race-page element
        WebDriverWait(driver, DEFAULT_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
        )
        time.sleep(2)

        # Venue / meeting title
        for sel in [
            "[class*='meeting-name']",
            "[class*='venue']",
            "[class*='MeetingName']",
            "h1",
        ]:
            el = safe_find(driver, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                race.venue = safe_text(el)
                break

        # Race number
        for sel in [
            "[class*='race-number']",
            "[class*='RaceNumber']",
            "[class*='raceNumber']",
        ]:
            el = safe_find(driver, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                race.race_number = safe_text(el)
                break

        # Race name / title
        for sel in [
            "[class*='race-name']",
            "[class*='RaceName']",
            "[class*='raceName']",
            "h2",
        ]:
            el = safe_find(driver, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                race.race_name = safe_text(el)
                break

        # Distance
        for sel in ["[class*='distance']", "[class*='Distance']"]:
            el = safe_find(driver, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                race.distance = safe_text(el)
                break

        # Track condition
        for sel in [
            "[class*='track-condition']",
            "[class*='TrackCondition']",
            "[class*='condition']",
        ]:
            el = safe_find(driver, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                race.track_condition = safe_text(el)
                break

        # Prize money
        for sel in [
            "[class*='prize-money']",
            "[class*='PrizeMoney']",
            "[class*='prize']",
        ]:
            el = safe_find(driver, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                race.prize_money = safe_text(el)
                break

        # Date / time - look in page title or meta elements
        for sel in ["[class*='date']", "[class*='time']", "time"]:
            el = safe_find(driver, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                text = safe_text(el)
                if not race.date:
                    race.date = text
                break

        log.info(
            "Race metadata — venue: %s, race: %s, name: %s, dist: %s",
            race.venue,
            race.race_number,
            race.race_name,
            race.distance,
        )
    except Exception as exc:
        log.warning("Could not fully parse race header: %s", exc)

    # --- Runners table ---
    try:
        runners = _scrape_tab_runners(driver)
        race.runners = runners
        log.info("Found %d runners", len(runners))
    except Exception as exc:
        log.error("Failed to scrape runners: %s", exc)

    return race


def _scrape_tab_runners(driver: webdriver.Remote) -> list[Runner]:
    runners: list[Runner] = []

    # TAB renders runners as rows inside a table or list.
    # We try several known CSS patterns.
    row_selectors = [
        "[class*='runner-row']",
        "[class*='RunnerRow']",
        "[class*='runner_row']",
        "tr[class*='runner']",
        "[class*='race-runner']",
        "[data-testid*='runner']",
        "[class*='competitor']",
        "tbody tr",
    ]

    rows = []
    for sel in row_selectors:
        rows = safe_find_all(driver, By.CSS_SELECTOR, sel)
        if rows:
            log.info("Runner rows found with selector: %s (%d rows)", sel, len(rows))
            break

    if not rows:
        log.warning("No runner rows found — attempting full page parse")
        # Dump page source for debugging
        return _parse_runners_from_page_source(driver)

    for row in rows:
        runner = Runner()

        # Runner number
        for sel in [
            "[class*='number']",
            "[class*='Number']",
            "td:first-child",
            "[class*='silk']",
        ]:
            el = safe_find(row, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                runner.number = safe_text(el)
                break

        # Horse name
        for sel in [
            "[class*='horse-name']",
            "[class*='HorseName']",
            "[class*='runner-name']",
            "[class*='RunnerName']",
            "[class*='name']",
            "a",
        ]:
            el = safe_find(row, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                runner.name = safe_text(el)
                break

        # Barrier
        for sel in ["[class*='barrier']", "[class*='Barrier']", "[class*='gate']"]:
            el = safe_find(row, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                runner.barrier = safe_text(el)
                break

        # Jockey
        for sel in [
            "[class*='jockey']",
            "[class*='Jockey']",
            "[class*='rider']",
        ]:
            el = safe_find(row, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                runner.jockey = safe_text(el)
                break

        # Trainer
        for sel in [
            "[class*='trainer']",
            "[class*='Trainer']",
        ]:
            el = safe_find(row, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                runner.trainer = safe_text(el)
                break

        # Weight
        for sel in [
            "[class*='weight']",
            "[class*='Weight']",
            "[class*='carried']",
        ]:
            el = safe_find(row, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                runner.weight = safe_text(el)
                break

        # Form
        for sel in [
            "[class*='form']",
            "[class*='Form']",
        ]:
            el = safe_find(row, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                runner.form = safe_text(el)
                break

        # Win odds
        for sel in [
            "[class*='win-price']",
            "[class*='WinPrice']",
            "[class*='win-odds']",
            "[class*='WinOdds']",
            "[class*='fixed-win']",
            "[class*='price']:first-child",
        ]:
            el = safe_find(row, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                runner.win_odds = safe_text(el)
                break

        # Place odds
        for sel in [
            "[class*='place-price']",
            "[class*='PlacePrice']",
            "[class*='place-odds']",
            "[class*='PlaceOdds']",
            "[class*='fixed-place']",
        ]:
            el = safe_find(row, By.CSS_SELECTOR, sel)
            if el and safe_text(el):
                runner.place_odds = safe_text(el)
                break

        # Scratched indicator
        scratched_el = safe_find(row, By.CSS_SELECTOR, "[class*='scratch']")
        if scratched_el:
            runner.scratched = True

        # Only add if we at least have a name
        if runner.name:
            runners.append(runner)
        else:
            # Log the raw text of the row for debugging
            raw = safe_text(row)
            if raw:
                log.debug("Skipping row with no name — raw text: %s", raw[:100])

    return runners


def _parse_runners_from_page_source(driver: webdriver.Remote) -> list[Runner]:
    """Fallback: parse visible text looking for runner patterns."""
    from selenium.webdriver.common.by import By

    runners: list[Runner] = []
    # Try to get all text content and look for numbered runner patterns
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text
        lines = [l.strip() for l in body_text.splitlines() if l.strip()]
        log.debug("Page has %d non-empty lines", len(lines))
        # Simple heuristic: lines starting with a digit followed by a likely horse name
        import re
        runner_pattern = re.compile(r"^(\d{1,2})\s+([A-Z][A-Z\s'()-]+)$")
        for line in lines:
            m = runner_pattern.match(line)
            if m:
                r = Runner(number=m.group(1), name=m.group(2).strip())
                runners.append(r)
    except Exception as exc:
        log.error("Fallback page parse failed: %s", exc)
    return runners


# ---------------------------------------------------------------------------
# RacingZone scraper
# ---------------------------------------------------------------------------

def scrape_racingzone_horse(driver: webdriver.Remote, horse_name: str) -> HorseStats:
    stats = HorseStats(name=horse_name)
    log.info("Searching RacingZone for: %s", horse_name)

    try:
        driver.get(RACINGZONE_HORSES_URL)
        time.sleep(2)

        # Find the search input — try several selectors
        search_input = None
        for sel in [
            "input[name='horse_name']",
            "input[name='name']",
            "input[placeholder*='horse' i]",
            "input[placeholder*='Horse' i]",
            "input[placeholder*='Find' i]",
            "input[type='search']",
            "input[type='text']",
            "#horse_name",
            "#name",
            "form input[type='text']",
        ]:
            search_input = safe_find(driver, By.CSS_SELECTOR, sel)
            if search_input:
                log.debug("Search input found with: %s", sel)
                break

        if not search_input:
            log.warning("Could not find search input on RacingZone for %s", horse_name)
            stats.error = "Search input not found"
            return stats

        # Clear and type the horse name
        search_input.clear()
        search_input.send_keys(horse_name)
        time.sleep(0.5)

        # Find and click search button
        search_btn = None
        for sel in [
            "button[type='submit']",
            "input[type='submit']",
            "button[class*='search']",
            "button[class*='Search']",
            "form button",
            "[class*='search-btn']",
            "[class*='SearchBtn']",
        ]:
            search_btn = safe_find(driver, By.CSS_SELECTOR, sel)
            if search_btn:
                log.debug("Search button found with: %s", sel)
                break

        if search_btn:
            search_btn.click()
        else:
            # Try pressing Enter on the input
            from selenium.webdriver.common.keys import Keys
            search_input.send_keys(Keys.RETURN)

        time.sleep(3)

        # Check if we landed on a results list or directly on a horse page
        current_url = driver.current_url
        log.debug("After search URL: %s", current_url)

        # If results list — pick first result
        result_links = safe_find_all(driver, By.CSS_SELECTOR, "[class*='result'] a")
        if not result_links:
            result_links = safe_find_all(driver, By.CSS_SELECTOR, "table a")
        if not result_links:
            # Try generic links in main content area
            result_links = safe_find_all(
                driver, By.CSS_SELECTOR, "main a, #content a, .content a"
            )

        if result_links:
            # Click the first result that contains the horse name (case-insensitive)
            clicked = False
            name_lower = horse_name.lower()
            for link in result_links:
                link_text = safe_text(link).lower()
                if name_lower in link_text or link_text in name_lower:
                    link.click()
                    clicked = True
                    time.sleep(3)
                    break
            if not clicked and result_links:
                result_links[0].click()
                time.sleep(3)

        # Now scrape the horse statistics page
        stats = _parse_horse_stats_page(driver, horse_name)

    except TimeoutException:
        log.warning("Timeout searching RacingZone for %s", horse_name)
        stats.error = "Timeout"
    except WebDriverException as exc:
        log.error("WebDriver error for %s: %s", horse_name, exc)
        stats.error = str(exc)[:200]

    return stats


def _parse_horse_stats_page(driver: webdriver.Remote, horse_name: str) -> HorseStats:
    stats = HorseStats(name=horse_name)

    try:
        page_text = driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        page_text = ""

    if not page_text.strip():
        stats.error = "Empty page"
        return stats

    # --- Horse profile info ---
    for sel in [
        "[class*='horse-name']",
        "[class*='HorseName']",
        "h1",
        "h2",
    ]:
        el = safe_find(driver, By.CSS_SELECTOR, sel)
        if el and safe_text(el):
            stats.name = safe_text(el)
            break

    # Sire / Dam / Colour / Sex / Age / Country — often in a detail table
    detail_labels = {
        "sire": ["sire", "father"],
        "dam": ["dam", "mother"],
        "colour": ["colour", "color"],
        "sex": ["sex", "gender"],
        "age": ["age"],
        "country": ["country", "origin"],
        "owner": ["owner"],
        "breeder": ["breeder"],
    }

    # Try structured table approach
    tables = safe_find_all(driver, By.CSS_SELECTOR, "table")
    for table in tables:
        rows = safe_find_all(table, By.CSS_SELECTOR, "tr")
        for row in rows:
            cells = safe_find_all(row, By.CSS_SELECTOR, "td, th")
            if len(cells) >= 2:
                label = safe_text(cells[0]).lower().rstrip(":")
                value = safe_text(cells[1])
                for attr, keywords in detail_labels.items():
                    if any(kw in label for kw in keywords):
                        setattr(stats, attr, value)

    # Try definition list approach (dl/dt/dd)
    dts = safe_find_all(driver, By.CSS_SELECTOR, "dt")
    dds = safe_find_all(driver, By.CSS_SELECTOR, "dd")
    for dt, dd in zip(dts, dds):
        label = safe_text(dt).lower().rstrip(":")
        value = safe_text(dd)
        for attr, keywords in detail_labels.items():
            if any(kw in label for kw in keywords):
                setattr(stats, attr, value)

    # --- Career statistics ---
    # Look for a career stats section
    career_section = None
    for sel in [
        "[class*='career']",
        "[class*='Career']",
        "#career",
        "[id*='career']",
    ]:
        career_section = safe_find(driver, By.CSS_SELECTOR, sel)
        if career_section:
            break

    stats_source = career_section if career_section else driver

    # Stats table rows: Starts | Wins | 2nd | 3rd | Win% | Place%
    stat_tables = safe_find_all(stats_source, By.CSS_SELECTOR, "table")
    for tbl in stat_tables:
        headers = [safe_text(h).lower() for h in safe_find_all(tbl, By.CSS_SELECTOR, "th")]
        data_rows = safe_find_all(tbl, By.CSS_SELECTOR, "tbody tr")

        if not data_rows:
            data_rows = safe_find_all(tbl, By.CSS_SELECTOR, "tr")[1:]

        for row in data_rows:
            cells = [safe_text(c) for c in safe_find_all(row, By.CSS_SELECTOR, "td")]
            if not cells:
                continue

            row_label = cells[0].lower() if cells else ""

            # Career row
            if "career" in row_label or "total" in row_label or "all" in row_label:
                _fill_stats(stats, cells, headers, prefix="career_")
            # Last 12 months
            elif "12" in row_label or "l12" in row_label or "year" in row_label:
                _fill_stats(stats, cells, headers, prefix="l12m_")

    # If we still have no career stats, try positional parsing
    if not stats.career_starts:
        _positional_stats_parse(stats, stat_tables)

    # --- Stats by category (distance, condition, track type, jockey, trainer) ---
    # Capture each section as raw text for the spreadsheet
    section_map = {
        "stats_by_distance": ["distance", "dist"],
        "stats_by_condition": ["condition", "going", "track cond"],
        "stats_by_track_type": ["track type", "surface"],
        "stats_by_jockey": ["jockey"],
        "stats_by_trainer": ["trainer"],
    }

    all_sections = safe_find_all(
        driver,
        By.CSS_SELECTOR,
        "section, [class*='section'], [class*='stats-group'], [class*='StatsGroup']",
    )
    for section in all_sections:
        section_text = safe_text(section)
        section_lower = section_text.lower()
        for attr, keywords in section_map.items():
            if any(kw in section_lower for kw in keywords):
                current = getattr(stats, attr)
                if not current:
                    setattr(stats, attr, section_text[:2000])

    # Fallback: capture full page text sections
    if not any([stats.stats_by_distance, stats.stats_by_condition]):
        _extract_sections_from_text(stats, page_text)

    log.info(
        "Parsed stats for %s — career: %s/%s/%s/%s",
        horse_name,
        stats.career_starts,
        stats.career_wins,
        stats.career_seconds,
        stats.career_thirds,
    )
    return stats


def _fill_stats(stats: HorseStats, cells: list, headers: list, prefix: str):
    """Map table cells to HorseStats fields using header names."""
    mapping = {
        "start": f"{prefix}starts",
        "win": f"{prefix}wins",
        "2nd": f"{prefix}seconds",
        "second": f"{prefix}seconds",
        "3rd": f"{prefix}thirds",
        "third": f"{prefix}thirds",
        "win%": f"{prefix}win_pct",
        "place%": f"{prefix}place_pct",
        "prize": f"{prefix}prize_money",
        "earning": f"{prefix}prize_money",
    }
    for i, cell_val in enumerate(cells[1:], start=1):
        if i <= len(headers):
            hdr = headers[i - 1].lower() if i - 1 < len(headers) else ""
        else:
            hdr = ""
        for key, attr in mapping.items():
            if key in hdr and hasattr(stats, attr):
                setattr(stats, attr, cell_val)
                break


def _positional_stats_parse(stats: HorseStats, tables: list):
    """Try to read stats from first table with enough columns."""
    for tbl in tables:
        rows = safe_find_all(tbl, By.CSS_SELECTOR, "tr")
        for row in rows:
            cells = [safe_text(c) for c in safe_find_all(row, By.CSS_SELECTOR, "td")]
            if len(cells) >= 4 and cells[0].isdigit():
                stats.career_starts = cells[0]
                stats.career_wins = cells[1] if len(cells) > 1 else ""
                stats.career_seconds = cells[2] if len(cells) > 2 else ""
                stats.career_thirds = cells[3] if len(cells) > 3 else ""
                return


def _extract_sections_from_text(stats: HorseStats, page_text: str):
    """Very basic section extraction from plain page text."""
    import re
    lines = page_text.splitlines()
    sections: dict[str, list] = {
        "distance": [],
        "condition": [],
        "jockey": [],
        "trainer": [],
    }
    current = None
    for line in lines:
        ll = line.lower().strip()
        if "distance" in ll:
            current = "distance"
        elif "condition" in ll or "going" in ll:
            current = "condition"
        elif "jockey" in ll:
            current = "jockey"
        elif "trainer" in ll:
            current = "trainer"
        elif current:
            sections[current].append(line)

    if sections["distance"]:
        stats.stats_by_distance = "\n".join(sections["distance"][:20])
    if sections["condition"]:
        stats.stats_by_condition = "\n".join(sections["condition"][:20])
    if sections["jockey"]:
        stats.stats_by_jockey = "\n".join(sections["jockey"][:20])
    if sections["trainer"]:
        stats.stats_by_trainer = "\n".join(sections["trainer"][:20])


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def save_to_excel(race: RaceInfo, horse_stats: list[HorseStats], output_path: str):
    output_dir = Path(output_path)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Build a safe filename from venue + race number
    safe_venue = "".join(c if c.isalnum() or c in "-_ " else "_" for c in race.venue)
    filename = f"{safe_venue}_R{race.race_number}.xlsx".replace(" ", "_")
    filepath = output_dir / filename

    log.info("Saving Excel file: %s", filepath)

    # --- Sheet 1: Race Overview ---
    race_data = {
        "Field": [
            "Venue",
            "Race Number",
            "Race Name",
            "Date",
            "Time",
            "Distance",
            "Track Condition",
            "Prize Money",
            "Race Class",
        ],
        "Value": [
            race.venue,
            race.race_number,
            race.race_name,
            race.date,
            race.time,
            race.distance,
            race.track_condition,
            race.prize_money,
            race.race_class,
        ],
    }
    df_race = pd.DataFrame(race_data)

    # --- Sheet 2: Runners (TAB data) ---
    runners_data = []
    for r in race.runners:
        runners_data.append(
            {
                "Number": r.number,
                "Horse Name": r.name,
                "Barrier": r.barrier,
                "Jockey": r.jockey,
                "Trainer": r.trainer,
                "Weight (kg)": r.weight,
                "Age": r.age,
                "Sex": r.sex,
                "Form": r.form,
                "Win Odds": r.win_odds,
                "Place Odds": r.place_odds,
                "Scratched": "Yes" if r.scratched else "No",
            }
        )
    df_runners = pd.DataFrame(runners_data)

    # --- Sheet 3: Horse Statistics (RacingZone data) ---
    stats_data = []
    for s in horse_stats:
        stats_data.append(
            {
                "Horse Name": s.name,
                "Sire": s.sire,
                "Dam": s.dam,
                "Colour": s.colour,
                "Sex": s.sex,
                "Age": s.age,
                "Country": s.country,
                "Owner": s.owner,
                "Breeder": s.breeder,
                "Career Starts": s.career_starts,
                "Career Wins": s.career_wins,
                "Career 2nds": s.career_seconds,
                "Career 3rds": s.career_thirds,
                "Career Win %": s.career_win_pct,
                "Career Place %": s.career_place_pct,
                "Career Prize Money": s.career_prize_money,
                "L12M Starts": s.l12m_starts,
                "L12M Wins": s.l12m_wins,
                "L12M 2nds": s.l12m_seconds,
                "L12M 3rds": s.l12m_thirds,
                "Stats by Distance": s.stats_by_distance,
                "Stats by Condition": s.stats_by_condition,
                "Stats by Track Type": s.stats_by_track_type,
                "Stats by Jockey": s.stats_by_jockey,
                "Stats by Trainer": s.stats_by_trainer,
                "Error": s.error,
            }
        )
    df_stats = pd.DataFrame(stats_data)

    # --- Sheet 4: Combined view ---
    if runners_data and stats_data:
        df_combined = df_runners.merge(
            df_stats,
            left_on="Horse Name",
            right_on="Horse Name",
            how="left",
            suffixes=("_TAB", "_RZ"),
        )
    else:
        df_combined = df_runners.copy()

    # Write to Excel with multiple sheets and some formatting
    with pd.ExcelWriter(str(filepath), engine="openpyxl") as writer:
        df_race.to_excel(writer, sheet_name="Race Info", index=False)
        df_runners.to_excel(writer, sheet_name="Runners (TAB)", index=False)
        df_stats.to_excel(writer, sheet_name="Horse Stats (RacingZone)", index=False)
        df_combined.to_excel(writer, sheet_name="Combined", index=False)

        # Auto-fit column widths
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        val = str(cell.value) if cell.value is not None else ""
                        # Cap width for long text fields
                        max_len = max(max_len, min(len(val), 60))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = max(max_len + 2, 10)

    log.info("Excel saved: %s", filepath)
    return str(filepath)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def resolve_output_path(raw: str) -> str:
    """Convert Windows-style path to a usable path on any OS."""
    if platform.system() == "Windows":
        return raw
    # On Linux/Mac, map C:\tab\scrape → ~/tab/scrape
    if raw.startswith("C:\\") or raw.startswith("c:\\"):
        rest = raw[3:].replace("\\", "/")
        return str(Path.home() / rest)
    return raw.replace("\\", "/")


def prompt_for_url() -> str:
    print("\n" + "=" * 60)
    print("  Horse Racing Scraper")
    print("=" * 60)
    print("Paste the full URL of the TAB race page and press Enter.")
    print("Example: https://www.tab.com.au/racing/meetings/RANDWICK/...")
    print()
    while True:
        url = input("TAB race URL: ").strip()
        if url.startswith("http"):
            return url
        print("  Please enter a valid URL starting with http.")


def main():
    parser = argparse.ArgumentParser(
        description="Scrape TAB race + RacingZone horse stats → Excel"
    )
    parser.add_argument(
        "--output",
        default=r"C:\tab\scrape",
        help=r"Output directory (default: C:\tab\scrape)",
    )
    parser.add_argument(
        "--browser",
        choices=["chrome", "firefox"],
        default="chrome",
        help="Browser to use (default: chrome)",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        default=True,
        help="Run browser in headless mode (default: True)",
    )
    parser.add_argument(
        "--no-headless",
        dest="headless",
        action="store_false",
        help="Show browser window",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=2.0,
        help="Seconds to wait between RacingZone requests (default: 2)",
    )
    args = parser.parse_args()

    url = prompt_for_url()

    output_path = resolve_output_path(args.output)
    log.info("Output directory: %s", output_path)

    driver = None
    try:
        driver = build_driver(browser=args.browser, headless=args.headless)
        driver.set_page_load_timeout(60)

        # Step 1: Scrape TAB race page
        race = scrape_tab_page(driver, url)

        if not race.runners:
            log.error(
                "No runners found on TAB page. Please check the URL and try again."
            )
            log.error("Page title: %s", driver.title)
            sys.exit(1)

        log.info(
            "Found %d runners. Starting RacingZone lookups...", len(race.runners)
        )

        # Step 2: Scrape RacingZone for each horse
        horse_stats: list[HorseStats] = []
        for i, runner in enumerate(race.runners):
            if runner.scratched:
                log.info("Skipping scratched horse: %s", runner.name)
                horse_stats.append(HorseStats(name=runner.name, error="Scratched"))
                continue

            log.info(
                "[%d/%d] Looking up: %s", i + 1, len(race.runners), runner.name
            )
            stats = scrape_racingzone_horse(driver, runner.name)
            horse_stats.append(stats)

            # Polite delay between requests
            if i < len(race.runners) - 1:
                time.sleep(args.delay)

        # Step 3: Save to Excel
        output_file = save_to_excel(race, horse_stats, output_path)
        print(f"\nDone! Excel file saved to: {output_file}")

    except KeyboardInterrupt:
        log.info("Interrupted by user.")
    except Exception as exc:
        log.exception("Unexpected error: %s", exc)
        sys.exit(1)
    finally:
        if driver:
            driver.quit()


if __name__ == "__main__":
    main()
