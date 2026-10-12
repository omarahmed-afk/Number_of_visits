import os
import json
from pathlib import Path
import re
import time
import random
import logging
import pandas as pd

from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# =====================================================================
# LOGGING
# =====================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("webpt_appointments.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)


def step(msg: str):
    log.info(f"{'─' * 60}")                          
    log.info(f"  ▶  {msg}")
    log.info(f"{'─' * 60}")


def ok(msg: str):
    log.info(f"  ✔  {msg}")


def warn(msg: str):
    log.warning(f"  ⚠  {msg}")


def fail(msg: str):
    log.error(f"  ✖  {msg}")


# =====================================================================
# HUMAN-LIKE DELAY
# =====================================================================

def human_delay(min_sec: float = 1.5, max_sec: float = 4.0):
    delay = random.uniform(min_sec, max_sec)
    log.info(f"  ·  Waiting {delay:.1f}s ...")
    time.sleep(delay)


def smart_wait(driver, wait, by, value, timeout=10):
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
        return el
    except Exception:
        return None


def wait_for_page_stable(driver, timeout=8):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except Exception:
        pass


# =====================================================================
# SETTINGS
# =====================================================================

_settings_path = Path(__file__).with_name("local_settings.json")
_local_settings = json.loads(_settings_path.read_text(encoding="utf-8")) if _settings_path.exists() else {}

WEBPT_URL    = "https://app.webpt.com"
USERNAME = os.environ.get("WEBPT_USERNAME", _local_settings.get("WEBPT_USERNAME", ""))
PASSWORD = os.environ.get("WEBPT_PASSWORD", _local_settings.get("WEBPT_PASSWORD", ""))
OUTPUT_PATH  = r"D:\Script (Previous Appointments)\Script (PIU)\results.xlsx"
WAIT_TIMEOUT = 20


# =====================================================================
# NAME / DATE HELPERS
# =====================================================================

def get_first_name(full_name: str) -> str:
    full_name = full_name.strip()
    full_name = re.sub(r'[,،‚]', ',', full_name)
    if "," in full_name:
        return full_name.split(",", 1)[1].strip()
    parts = full_name.split()
    return parts[-1] if parts else full_name


def _get_last_name(full_name: str) -> str:
    full_name = full_name.strip()
    full_name = re.sub(r'[,،‚]', ',', full_name)
    if "," in full_name:
        return full_name.split(",", 1)[0].strip()
    parts = full_name.split()
    return " ".join(parts[:-1]) if len(parts) > 1 else full_name


def normalize_date(date_str: str) -> str:
    try:
        clean = re.sub(r'\(.*?\)', '', date_str.strip())
        clean = re.sub(r'(\d+)(st|nd|rd|th)', r'\1', clean).strip()
        return pd.to_datetime(clean).strftime("%m/%d/%Y")
    except Exception:
        return date_str.strip()


# =====================================================================
# DRIVER
# =====================================================================

def create_driver() -> webdriver.Chrome:
    step("Launching Chrome browser")
    options = webdriver.ChromeOptions()

    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    options.add_argument("--start-maximized")

    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    ]
    options.add_argument(f"--user-agent={random.choice(user_agents)}")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
            Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3] });
            Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
        """
    })

    ok("Browser launched successfully")
    return driver


# =====================================================================
# LOGIN
# =====================================================================

def login(driver, wait):
    step("Logging in to WebPT")

    driver.get(WEBPT_URL)
    wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(USERNAME)
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

    wait.until(EC.presence_of_element_located((By.ID, "password"))).send_keys(PASSWORD)
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

    for prompt_text in ["Yes, oust them!", "Yes"]:
        try:
            btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//button[contains(text(),'{prompt_text}')]")
                )
            )
            btn.click()
            ok(f"Dismissed prompt: '{prompt_text}'")
            break
        except Exception:
            pass

    wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Advanced Search")))
    ok("Login successful — dashboard ready")

    # ── Open View Schedule once, right after login ──────────────────
    # From this point on, the script stays on the View Schedule page
    # for the entire run. It does NOT return to the dashboard between
    # patients. The flow is:
    #   View Schedule → Clinic → Date → Search → Close popup
    #   → (next patient) Clinic (skipped if same) → Date → Search → ...
    step("Opening View Schedule (one-time)")
    _open_view_schedule(driver, wait)


def _open_view_schedule(driver, wait):
    """
    Opens the View Schedule page from the current page.
    Tries multiple strategies. Verifies success by checking for the
    ClinicChange dropdown.
    """
    for attempt in range(1, 4):
        log.info(f"  →  Opening View Schedule (attempt {attempt}/3) ...")

        clicked = False
        for locator in [
            (By.LINK_TEXT,          "View Schedule"),
            (By.PARTIAL_LINK_TEXT,  "View Schedule"),
            (By.XPATH,              "//a[contains(text(),'View Schedule')]"),
            (By.XPATH,              "//*[contains(text(),'View Schedule')]"),
            (By.CSS_SELECTOR,       "a[href*='scheduler']"),
        ]:
            try:
                el = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(locator))
                driver.execute_script("arguments[0].click();", el)
                wait_for_page_stable(driver, timeout=20)
                human_delay(1.0, 2.5)
                clicked = True
                break
            except Exception:
                continue

        # Check if View Schedule is now open
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ClinicChange"))
            )
            ok(f"View Schedule opened (attempt {attempt})")
            return True
        except Exception:
            pass

        # Fallback: direct URL
        try:
            log.info(f"  →  Trying direct URL ...")
            driver.get(WEBPT_URL + "/scheduler/")
            wait_for_page_stable(driver, timeout=20)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ClinicChange"))
            )
            ok(f"View Schedule opened via direct URL (attempt {attempt})")
            return True
        except Exception:
            pass

        if attempt < 3:
            log.info(f"  →  Will retry ...")
            time.sleep(3)

    warn("Could not open View Schedule after 3 attempts")
    return False


# =====================================================================
# CLINIC SELECTION
# =====================================================================

_current_clinic = None

# Tracks whether the No Show + Cancelled filters are currently active.
# When they are, we can reuse them for the next patient (no need to
# toggle them off and on again). They reset whenever the clinic changes.
_extra_filters_active = False


def _get_clinic_select_element(driver):
    try:
        return driver.find_element(By.ID, "ClinicChange")
    except Exception:
        try:
            return driver.find_element(By.NAME, "ClinicChange")
        except Exception:
            return None


def _read_current_clinic_from_page(driver) -> str:
    try:
        sel_el = _get_clinic_select_element(driver)
        if sel_el is None:
            return ""
        try:
            sel = Select(sel_el)
            first_selected = sel.first_selected_option
            text = first_selected.text.strip()
            if text:
                return text
        except Exception:
            pass
        try:
            text = driver.execute_script(
                "var s=arguments[0]; "
                "if(s.selectedIndex<0) return ''; "
                "return s.options[s.selectedIndex].text || '';",
                sel_el
            )
            return (text or "").strip()
        except Exception:
            pass
    except Exception:
        pass
    return ""


def select_clinic_in_schedule(driver, clinic_name: str):
    global _current_clinic, _extra_filters_active

    target_clinic = clinic_name.strip()
    clinic_lower  = target_clinic.lower()

    if _current_clinic and _current_clinic.lower() == clinic_lower:
        ok(f"Clinic '{target_clinic}' already selected (cached) — skipping")
        return

    current_on_page = _read_current_clinic_from_page(driver)
    log.info(f"  →  Currently selected clinic: '{current_on_page}'")
    if current_on_page and (
        current_on_page.lower() == clinic_lower
        or clinic_lower in current_on_page.lower()
        or current_on_page.lower() in clinic_lower
    ):
        ok(f"Clinic '{current_on_page}' already shown — skipping")
        _current_clinic = current_on_page
        return

    # Clinic is going to CHANGE — invalidate the extra-filters state
    # because the new clinic's view will reset all filter selections.
    if _extra_filters_active:
        log.info("  →  Clinic is changing — filters will reset")
        _extra_filters_active = False

    log.info(f"  →  Changing clinic to '{target_clinic}' via visual dropdown ...")

    trigger_clicked = False

    try:
        sel_el = _get_clinic_select_element(driver)
        candidates = driver.find_elements(
            By.XPATH,
            "//input[@type='text' or not(@type)]"
        )
        for cand in candidates:
            try:
                if not cand.is_displayed():
                    continue
                val = (cand.get_attribute("value") or "").strip()
                if current_on_page and val == current_on_page:
                    log.info(f"  →  Found visible clinic input (value='{val}') — clicking ...")
                    try:
                        cand.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", cand)
                    time.sleep(1)
                    trigger_clicked = True
                    break
            except Exception:
                continue
    except Exception as e:
        log.debug(f"Strategy A error: {e}")

    if not trigger_clicked:
        try:
            for sel_str in [
                "img[class*='trigger']",
                "div[class*='trigger']",
                "img[class*='arrow']",
                "div[class*='arrow-trigger']",
                "*[class*='x-form-trigger']",
            ]:
                triggers = driver.find_elements(By.CSS_SELECTOR, sel_str)
                for trg in triggers:
                    try:
                        if not trg.is_displayed():
                            continue
                        log.info(f"  →  Trying trigger '{sel_str}' ...")
                        try:
                            trg.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", trg)
                        time.sleep(1)
                        trigger_clicked = True
                        break
                    except Exception:
                        continue
                if trigger_clicked:
                    break
        except Exception as e:
            log.debug(f"Strategy B error: {e}")

    if not trigger_clicked:
        try:
            sel_el = _get_clinic_select_element(driver)
            if sel_el:
                driver.execute_script(
                    """
                    var s = arguments[0];
                    var p = s.parentElement;
                    if (p) {
                        var rect = p.getBoundingClientRect();
                        var ev = new MouseEvent('click', {
                            bubbles: true, cancelable: true, view: window,
                            clientX: rect.right - 10,
                            clientY: rect.top + rect.height / 2
                        });
                        p.dispatchEvent(ev);
                    }
                    """,
                    sel_el
                )
                time.sleep(1)
                trigger_clicked = True
                log.info("  →  Clicked dropdown via parent element")
        except Exception as e:
            log.debug(f"Strategy C error: {e}")

    if not trigger_clicked:
        warn("Could not click clinic dropdown trigger")
        return

    time.sleep(0.8)

    search_typed = False
    try:
        all_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], input:not([type])")
        for inp in all_inputs:
            try:
                if not inp.is_displayed():
                    continue
                val = (inp.get_attribute("value") or "").strip()
                if val:
                    continue
                cls = (inp.get_attribute("class") or "").lower()
                iid  = (inp.get_attribute("id") or "").lower()
                name = (inp.get_attribute("name") or "").lower()
                if "patient" in iid or "patient" in name or "date" in iid:
                    continue
                log.info(f"  →  Typing '{target_clinic}' into search box (id='{iid}', class='{cls[:40]}')")
                inp.clear()
                time.sleep(0.3)
                inp.send_keys(target_clinic)
                time.sleep(1.2)
                search_typed = True
                break
            except Exception:
                continue
    except Exception as e:
        log.debug(f"Search type error: {e}")

    if not search_typed:
        log.info("  →  No search box found; will scan the visible list directly")

    matched = False
    try:
        item_selectors = [
            "div.x-combo-list-item",
            "div.x-boundlist-item",
            "li.x-boundlist-item",
            "div[class*='combo-list-item']",
            "div[class*='boundlist-item']",
            "div[class*='list-item']",
        ]
        all_items = []
        for isel in item_selectors:
            found = driver.find_elements(By.CSS_SELECTOR, isel)
            for f in found:
                if f.is_displayed():
                    all_items.append(f)

        log.info(f"  →  Visible list items: {len(all_items)}")

        item_texts = []
        for it in all_items[:25]:
            try:
                t = it.text.strip()
                if t:
                    item_texts.append(t)
            except Exception:
                pass
        log.info(f"  →  Visible items (first 25): {item_texts}")

        for it in all_items:
            try:
                t = it.text.strip()
                if t.lower() == clinic_lower:
                    log.info(f"  →  Clicking exact match: '{t}'")
                    driver.execute_script("arguments[0].click();", it)
                    time.sleep(2.5)
                    wait_for_page_stable(driver, timeout=10)
                    _current_clinic = t
                    ok(f"Clinic selected: '{t}'")
                    matched = True
                    return
            except Exception:
                continue

        if not matched:
            for it in all_items:
                try:
                    t = it.text.strip()
                    if not t:
                        continue
                    if clinic_lower in t.lower() or t.lower() in clinic_lower:
                        log.info(f"  →  Clicking substring match: '{t}'")
                        driver.execute_script("arguments[0].click();", it)
                        time.sleep(2.5)
                        wait_for_page_stable(driver, timeout=10)
                        _current_clinic = t
                        ok(f"Clinic selected: '{t}'")
                        matched = True
                        return
                except Exception:
                    continue
    except Exception as e:
        log.debug(f"Item click error: {e}")

    if not matched:
        log.info("  →  Visual dropdown approach didn't match; trying hidden <select> fallback")
        try:
            sel_el = _get_clinic_select_element(driver)
            if sel_el is None:
                warn("ClinicChange <select> not found")
                return

            js_options = driver.execute_script(
                "var s=arguments[0],r=[];for(var i=0;i<s.options.length;i++){var o=s.options[i];r.push({v:o.value,t:o.text});}return r;",
                sel_el
            ) or []

            log.info(f"  →  Hidden <select> has {len(js_options)} options")

            matched_value = None
            matched_text  = None
            for o in js_options:
                if (o.get('t') or '').strip().lower() == clinic_lower:
                    matched_value = o['v']; matched_text = o['t'].strip(); break
            if not matched_value:
                for o in js_options:
                    t = (o.get('t') or '').strip().lower()
                    if t and (clinic_lower in t or t in clinic_lower):
                        matched_value = o['v']; matched_text = o['t'].strip(); break

            if matched_value:
                driver.execute_script(
                    "var s=arguments[0],v=arguments[1];"
                    "for(var i=0;i<s.options.length;i++){if(s.options[i].value===v){s.selectedIndex=i;break;}}"
                    "s.value=v;"
                    "var ev;try{ev=new Event('change',{bubbles:true,cancelable:true});}"
                    "catch(e){ev=document.createEvent('HTMLEvents');ev.initEvent('change',true,true);}"
                    "s.dispatchEvent(ev);",
                    sel_el, matched_value
                )
                time.sleep(3)
                wait_for_page_stable(driver, timeout=10)
                _current_clinic = matched_text
                ok(f"Clinic selected via hidden <select>: '{matched_text}'")
                return
        except Exception as e:
            warn(f"Hidden <select> fallback failed: {e}")

    warn(f"Could not select clinic '{clinic_name}' — using current")


# =====================================================================
# CALENDAR NAVIGATION
# =====================================================================

def navigate_schedule_to_date(driver, target_date: str):
    try:
        target_normalized = normalize_date(target_date)
        target_dt         = pd.to_datetime(target_normalized)
        target_day        = str(target_dt.day)
        target_month_abbr = target_dt.strftime("%b")
        target_year       = target_dt.strftime("%Y")
        target_month_full = target_dt.strftime("%b %Y")

        log.info(f"  →  Target: {target_month_full} day {target_day}")

        current = _read_calendar_header(driver)
        log.info(f"  →  Current calendar header: '{current}'")

        if current and target_month_abbr.lower() in current.lower() and target_year in current:
            ok(f"Already on correct month: {current}")
        else:
            reached = _calendar_select_month_year(driver, target_month_abbr, target_year)
            if not reached:
                warn(f"Arrow navigation failed — cannot safely navigate to {target_month_full}")
                return

        time.sleep(1)
        clicked = _calendar_click_day(driver, target_day, target_month_full)

        if not clicked:
            warn(f"Could not click day {target_day} in {target_month_full}")
            return

        time.sleep(2)

        verify_header = _read_calendar_header(driver)
        if verify_header and (target_month_abbr.lower() not in verify_header.lower() or target_year not in verify_header):
            warn(f"After clicking day {target_day}, calendar shows '{verify_header}' — expected '{target_month_full}'")
            warn("Attempting recovery via arrow navigation ...")
            recovered = _calendar_select_month_year(driver, target_month_abbr, target_year)
            if recovered:
                time.sleep(1)
                _calendar_click_day(driver, target_day, target_month_full)
                time.sleep(2)
                ok(f"Recovery successful: navigated to {target_normalized}")
            else:
                warn("Recovery failed — proceeding with best effort")
        else:
            ok(f"Schedule navigated to: {target_normalized}")

    except Exception as e:
        warn(f"navigate_schedule_to_date error: {e}")


def _read_calendar_header(driver) -> str:
    pattern = re.compile(
        r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}',
        re.IGNORECASE
    )

    try:
        tables = driver.find_elements(By.TAG_NAME, "table")
        for tbl in tables:
            tbl_text = tbl.text
            if "Su" in tbl_text and "Mo" in tbl_text and "Sa" in tbl_text:
                rows = tbl.find_elements(By.TAG_NAME, "tr")
                if rows:
                    header_text = rows[0].text.strip()
                    m = pattern.search(header_text)
                    if m:
                        return m.group(0)
                for cell in tbl.find_elements(By.XPATH, ".//td | .//th"):
                    text = cell.text.strip()
                    if re.match(r'^\d{1,2}$', text):
                        continue
                    m = pattern.search(text)
                    if m:
                        return m.group(0)
    except Exception:
        pass

    for tag in ("td", "th"):
        try:
            elements = driver.find_elements(By.TAG_NAME, tag)
            for el in elements:
                text = el.text.strip()
                if re.match(r'^\d{1,2}$', text):
                    continue
                m = pattern.search(text)
                if m:
                    return m.group(0)
        except Exception:
            continue
    return ""


def _calendar_select_month_year(driver, target_month_str: str, target_year_str: str) -> bool:
    target_month_full = f"{target_month_str} {target_year_str}"
    log.info(f"  →  Arrow-navigating calendar to: {target_month_full}")

    try:
        target_dt = pd.to_datetime(f"01 {target_month_str} {target_year_str}")
    except Exception as e:
        warn(f"Cannot parse target date '{target_month_full}': {e}")
        return False

    for attempt in range(24):
        time.sleep(1)
        current = _read_calendar_header(driver)
        log.info(f"  →  Calendar header (attempt {attempt+1}): '{current}'")

        if not current:
            warn("Cannot read calendar header")
            time.sleep(1.5)
            continue

        if target_month_str.lower() in current.lower() and target_year_str in current:
            ok(f"Correct month reached: {current}")
            return True

        try:
            cur_dt     = pd.to_datetime("01 " + current)
            go_forward = target_dt > cur_dt
        except Exception:
            go_forward = True

        if go_forward:
            _click_calendar_arrow(driver, ">")
        else:
            _click_calendar_arrow(driver, "<")

    warn(f"Could not reach {target_month_full} after 24 attempts")
    return False


def _click_calendar_arrow(driver, symbol: str):
    try:
        tables = driver.find_elements(By.TAG_NAME, "table")
        cal_table = None
        for tbl in tables:
            txt = tbl.text
            if "Mo" in txt and "Su" in txt and "Sa" in txt:
                cal_table = tbl
                break

        if cal_table:
            for a in cal_table.find_elements(By.TAG_NAME, "a"):
                try:
                    t = a.text.strip()
                    if t == symbol and chr(10) not in a.text:
                        clicked_ok = False
                        try:
                            a.click()
                            clicked_ok = True
                            log.info(f"  →  Clicked '{symbol}' via native .click()")
                        except Exception:
                            pass
                        if not clicked_ok:
                            try:
                                from selenium.webdriver.common.action_chains import ActionChains
                                ActionChains(driver).move_to_element(a).click().perform()
                                clicked_ok = True
                                log.info(f"  →  Clicked '{symbol}' via ActionChains")
                            except Exception:
                                pass
                        if not clicked_ok:
                            try:
                                driver.execute_script("""
                                    var el = arguments[0];
                                    var e1 = new MouseEvent('mousedown', {bubbles:true, cancelable:true});
                                    var e2 = new MouseEvent('mouseup',   {bubbles:true, cancelable:true});
                                    var e3 = new MouseEvent('click',     {bubbles:true, cancelable:true});
                                    el.dispatchEvent(e1);
                                    el.dispatchEvent(e2);
                                    el.dispatchEvent(e3);
                                """, a)
                                clicked_ok = True
                                log.info(f"  →  Clicked '{symbol}' via MouseEvent dispatch")
                            except Exception:
                                pass
                        if clicked_ok:
                            return
                except Exception:
                    continue

            for td in cal_table.find_elements(By.TAG_NAME, "td"):
                try:
                    raw = td.text
                    stripped = raw.strip()
                    if stripped == symbol and chr(10) not in raw:
                        driver.execute_script("arguments[0].click();", td)
                        log.info(f"  →  Clicked '{symbol}' arrow via <td> in calendar table")
                        return
                except Exception:
                    continue

        for tag in ("a", "td", "span", "input", "button"):
            for el in driver.find_elements(By.TAG_NAME, tag):
                try:
                    raw = el.text
                    stripped = raw.strip()
                    val = (el.get_attribute("value") or "").strip()
                    if (stripped == symbol or val == symbol) and chr(10) not in raw:
                        driver.execute_script("arguments[0].click();", el)
                        log.info(f"  →  Clicked '{symbol}' arrow via <{tag}> (page scan)")
                        return
                except Exception:
                    continue

        warn(f"Arrow '{symbol}' not found in calendar")

        try:
            js = f"""
            var all = document.querySelectorAll('td, a, span, button, input');
            for (var i = 0; i < all.length; i++) {{
                var el = all[i];
                var t = (el.innerText || el.textContent || el.value || '').trim();
                if (t === '{symbol}' && !t.includes('\\n')) {{
                    el.click();
                    return true;
                }}
            }}
            return false;
            """
            result = driver.execute_script(js)
            if result:
                log.info(f"  →  Clicked '{symbol}' arrow via JavaScript")
            else:
                warn(f"  →  JavaScript arrow click also failed for '{symbol}'")
        except Exception as e2:
            warn(f"JavaScript arrow fallback failed: {e2}")

    except Exception as e:
        warn(f"_click_calendar_arrow('{symbol}') error: {e}")


def _calendar_click_day(driver, target_day: str, target_month_year: str = "") -> bool:
    try:
        cal_table = None
        tables = driver.find_elements(By.TAG_NAME, "table")
        for tbl in tables:
            tbl_text = tbl.text
            if "Su" in tbl_text and "Mo" in tbl_text and "Sa" in tbl_text:
                cal_table = tbl
                break

        if not cal_table:
            warn(f"Calendar table not found — cannot safely click day {target_day}")
            return False

        day_links = cal_table.find_elements(By.TAG_NAME, "a")
        log.info(f"  →  Calendar table has {len(day_links)} <a> links")

        candidates = []
        for link in day_links:
            try:
                link_text = link.text.strip()
                if link_text != target_day:
                    continue

                try:
                    parent_td = link.find_element(By.XPATH, "./ancestor::td[1]")
                    td_class = (parent_td.get_attribute("class") or "").lower()
                except Exception:
                    td_class = ""

                is_other_month = ("prevday" in td_class or "nextday" in td_class
                                   or "outside" in td_class or "other-month" in td_class
                                   or "x-datepicker-prevday" in td_class
                                   or "x-datepicker-nextday" in td_class)
                if is_other_month:
                    log.info(f"  →  Skipping day {target_day} from adjacent month (class='{td_class}')")
                    continue

                link_class = (link.get_attribute("class") or "").lower()
                if "prevday" in link_class or "nextday" in link_class:
                    log.info(f"  →  Skipping day {target_day} (link class='{link_class}')")
                    continue

                candidates.append((link, td_class, link_class))
            except Exception:
                continue

        log.info(f"  →  Found {len(candidates)} valid candidate(s) for day {target_day}")

        if not candidates:
            warn(f"Day {target_day} not found in current month inside the calendar table")
            for link in day_links:
                if link.text.strip() == target_day:
                    log.info(f"  →  Fallback: clicking first '{target_day}' link without class filter")
                    driver.execute_script("arguments[0].click();", link)
                    return True
            return False

        link, td_class, link_class = candidates[0]
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", link)
        ok(f"Clicked day {target_day} in calendar (active month)")
        return True

    except Exception as e:
        warn(f"_calendar_click_day error: {e}")
        return False


# =====================================================================
# PROCESS A SINGLE APPOINTMENT
# (Same flow as the original — but only extracts Status + Appt_Date)
# =====================================================================

def _process_single_appointment(driver, full_name: str, matched_el, icon_status: str):
    """
    Processes ONE matched appointment from the schedule:

    Logic:
      1. THE ICON IS THE PRIMARY SOURCE OF STATUS.
         The icon next to the patient's name (√, «, Φ, ?) decides the
         status BEFORE the popup is even opened.

      2. READ TREATING THERAPIST BEFORE CLICKING.
         The appointment's X-position tells us which therapist column
         it's under. This must be read before the popup opens (since
         the popup might obscure the schedule).

      3. THE POPUP LINK PROVIDES DATE/TIME.
         When the popup opens, we extract the date and time from the
         status link inside it.

      4. NO Scheduler History navigation.

    Returns (status, appt_date, therapist)
    """
    # The icon already gave us the status — that's our answer.
    final_status = icon_status if icon_status else "Other"
    appt_date    = ""
    therapist    = ""
    appt_type    = ""

    log.info(f"  →  Icon-detected status (primary): '{final_status}'")

    # ── Step 0a: Read the Treating Therapist BEFORE clicking ─────────
    # The popup might cover schedule columns once opened, so we must
    # read the therapist's name first while the appointment is still
    # visible in its column.
    try:
        therapist = _get_therapist_for_appointment(driver, matched_el)
        if therapist:
            ok(f"  →  Treating Therapist: '{therapist}'")
        else:
            log.info("  →  Treating Therapist: (could not determine)")
    except Exception as e:
        log.debug(f"Therapist read error: {e}")

    # ── Step 0b: Read the Appointment Type from tooltip (hover) ─────
    # Hover over the appointment to trigger its tooltip (e.g. "School INDV 9:30am"),
    # then take only the part before the time → "School INDV".
    # Must be done BEFORE clicking, since clicking opens the popup.
    try:
        appt_type = _get_tooltip_for_appointment(driver, matched_el)
        if appt_type:
            ok(f"  →  Appointment Type: '{appt_type}'")
        else:
            log.info("  →  Appointment Type: (could not determine)")
    except Exception as e:
        log.debug(f"Appointment type read error: {e}")

    # ── Step 1: Click the appointment to open the popup ─────────────
    try:
        driver.execute_script("arguments[0].scrollIntoView(true);", matched_el)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", matched_el)
        time.sleep(2)
    except Exception as e:
        log.debug(f"Click error: {e}")
        warn(f"Could not click appointment for '{full_name}'")
        return final_status, "", therapist, appt_type

    # ── Step 2: Wait briefly for popup to render ────────────────────
    time.sleep(1.2)

    # ── Step 3: Find the status link in the popup ───────────────────
    # The popup contains a link like:
    #   "Checked Out by Sandra Emad 04/29/2026 08:58 AM"
    #   "Checked In by Hagar Alaa 04/27/2026 09:49 AM"
    #   "No Show by Mike 04/22/2026 10:00 AM"
    #   "Cancelled by Sarah 04/22/2026 02:30 PM"
    #   "Edited by Elsayed Moussa 05/05/2026 02:13 PM"
    #   "Created by Hagar Alaa 05/09/2026 10:30 AM"
    popup_link_text = ""
    try:
        all_links = driver.find_elements(By.TAG_NAME, "a")
        for el in all_links:
            try:
                if not el.is_displayed():
                    continue
                link_href = (el.get_attribute("href") or "")
                # The status link points to scheduler/history with date=...&uid=...
                if ("scheduler/history" in link_href or "schedulerHistory" in link_href) \
                        and ("date=" in link_href or "uid=" in link_href):
                    popup_link_text = el.text.strip()
                    log.info(f"  →  Found popup link: '{popup_link_text[:150]}'")
                    break
            except Exception:
                continue

        # Fallback: any visible link whose text matches "<event> by <user> <date> <time>"
        if not popup_link_text:
            for el in all_links:
                try:
                    if not el.is_displayed():
                        continue
                    txt = el.text.strip()
                    if not txt:
                        continue
                    tl = txt.lower()
                    if (("checked out" in tl or "checked in" in tl or "no show" in tl
                         or "cancelled" in tl or "canceled" in tl
                         or "edited" in tl or "created" in tl)
                            and " by " in tl
                            and re.search(r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}', txt)):
                        popup_link_text = txt
                        log.info(f"  →  Found popup link (fallback): '{popup_link_text[:150]}'")
                        break
                except Exception:
                    continue
    except Exception as e:
        log.debug(f"Popup link search error: {e}")

    # ── Step 4: Extract Date + Time from the link (status comes from icon) ─
    # We do NOT change final_status based on the link's event word —
    # the icon is authoritative. We only take the date & time.
    if popup_link_text:
        try:
            # Try the strict pattern first:
            #   "<Event> by <User> MM/DD/YYYY HH:MM AM"
            m = re.match(
                r'^(Checked Out|Checked In|No Show|Cancelled|Edited|Created)\s+by\s+(.+?)\s+(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})\s+(\d{1,2}:\d{2}\s*[AP]M)\s*$',
                popup_link_text,
                re.IGNORECASE,
            )
            if m:
                event_word = m.group(1).strip()
                date_part  = m.group(3).strip()
                time_part  = m.group(4).strip()

                try:
                    appt_date = pd.to_datetime(f"{date_part} {time_part}").strftime("%Y-%m-%d %I:%M %p")
                except Exception:
                    appt_date = f"{date_part} {time_part}"

                log.info(f"  →  Link event word: '{event_word}' (ignored — icon decides status)")
                ok(f"  ✓ Date='{date_part}' | Time='{time_part}'")
            else:
                # Loose extraction: just find any date + time anywhere in the text
                log.info(f"  →  Strict pattern didn't match — trying loose extraction")
                dt_match = re.search(
                    r'(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})\s+(\d{1,2}:\d{2}\s*[AP]M)',
                    popup_link_text,
                    re.IGNORECASE
                )
                if dt_match:
                    date_part = dt_match.group(1).strip()
                    time_part = dt_match.group(2).strip()
                    try:
                        appt_date = pd.to_datetime(f"{date_part} {time_part}").strftime("%Y-%m-%d %I:%M %p")
                    except Exception:
                        appt_date = f"{date_part} {time_part}"
                    ok(f"  ✓ Loose extract: Date='{date_part}' | Time='{time_part}'")
                else:
                    log.info(f"  →  No date/time found in link — appt_date will be empty")
        except Exception as e:
            log.debug(f"  →  Link parse error: {e}")
    else:
        log.info("  →  No popup link found — appt_date will be empty")

    # ── Step 5: Close the popup and return to schedule ──────────────
    _ensure_on_schedule_view(driver)

    ok(f"FINAL: Status='{final_status}' | Date='{appt_date}' | Therapist='{therapist}' | Type='{appt_type}'")
    return final_status, appt_date, therapist, appt_type


def _close_popup_if_open(driver):
    from selenium.webdriver.common.keys import Keys

    try:
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys(Keys.ESCAPE)
        time.sleep(0.4)
    except Exception:
        pass

    close_selectors = [
        "div.x-tool-close",
        "img.x-tool-close",
        "div.x-tool-img.x-tool-close",
        "a.close-button",
        "div.close-button",
        "img.closeButton",
        "img[src*='close']",
        "*[class*='x-tool-close']",
        "*[title='Close']",
        "*[aria-label='Close']",
        "img[alt='close']",
        "img[alt='Close']",
    ]

    for sel in close_selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in elements:
                try:
                    if el.is_displayed():
                        driver.execute_script("arguments[0].click();", el)
                        time.sleep(0.5)
                        return
                except Exception:
                    continue
        except Exception:
            continue

    try:
        x_buttons = driver.find_elements(
            By.XPATH,
            "//*[normalize-space(text())='×' or normalize-space(text())='X' or normalize-space(text())='✕']"
        )
        for btn in x_buttons:
            try:
                if btn.is_displayed() and btn.size.get('width', 0) < 50:
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.5)
                    return
            except Exception:
                continue
    except Exception:
        pass


def _ensure_on_schedule_view(driver):
    try:
        if len(driver.window_handles) > 1:
            main_win = driver.window_handles[0]
            for handle in driver.window_handles[1:]:
                try:
                    driver.switch_to.window(handle)
                    driver.close()
                except Exception:
                    pass
            driver.switch_to.window(main_win)
            log.info("  →  Closed extra tabs, returned to schedule")
    except Exception as e:
        log.debug(f"Tab cleanup error: {e}")

    _close_popup_if_open(driver)
    time.sleep(0.5)


def _get_tooltip_for_appointment(driver, appt_element) -> str:
    """
    Hovers over the appointment to trigger its tooltip, then reads the
    tooltip text and returns the part BEFORE the time.

    WebPT tooltip HTML structure (confirmed):
      <div class="tooltiptext">
          <div>Teletherapy 6:30pm</div>     ← first child div = type + time
          <div>...patient name...</div>
          <div>(718)753-7116</div>
          ...
      </div>

    Strategy:
      1. Hover over the appointment
      2. Find the visible .tooltiptext element
      3. Read its FIRST child <div> only (e.g. "Teletherapy 6:30pm")
      4. Strip the time → returns "Teletherapy"

    Returns empty string if tooltip can't be read.
    """
    try:
        from selenium.webdriver.common.action_chains import ActionChains

        # ── Step 0: Clear any leftover tooltip from a previous appt ──
        try:
            body = driver.find_element(By.TAG_NAME, "body")
            ActionChains(driver).move_to_element_with_offset(body, 0, 0).perform()
            time.sleep(0.5)
        except Exception:
            pass

        # ── Step 1: Scroll appointment into view ────────────────────
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", appt_element)
            time.sleep(0.3)
        except Exception:
            pass

        # Get the appointment's position so we can match the NEAREST tooltip
        appt_box = driver.execute_script("""
            var r = arguments[0].getBoundingClientRect();
            return {left: r.left, right: r.right, top: r.top, bottom: r.bottom,
                    cx: (r.left+r.right)/2, cy: (r.top+r.bottom)/2};
        """, appt_element)
        appt_cx = appt_box.get('cx', 0) if appt_box else 0
        appt_cy = appt_box.get('cy', 0) if appt_box else 0

        # ── Step 2: Hover + read tooltip, with retries for "Loading" ─
        first_line = ""
        max_tries = 5
        for attempt in range(1, max_tries + 1):
            # Re-hover each attempt (keeps the tooltip alive / refreshes it)
            try:
                ActionChains(driver).move_to_element(appt_element).perform()
            except Exception:
                pass
            # Wait progressively longer; the tooltip loads its content async
            time.sleep(1.0 + attempt * 0.4)

            # Find the tooltip NEAREST to the appointment (by center distance).
            # This avoids reading a different appointment's tooltip.
            first_line = driver.execute_script("""
                var apptCx = arguments[0];
                var apptCy = arguments[1];
                var tips = document.querySelectorAll('.tooltiptext');
                var best = '';
                var bestDist = Infinity;
                for (var i = 0; i < tips.length; i++) {
                    var tip = tips[i];
                    var r = tip.getBoundingClientRect();
                    if (r.width < 10 || r.height < 5) continue;
                    var style = window.getComputedStyle(tip);
                    if (style.display === 'none' || style.visibility === 'hidden') continue;
                    if (parseFloat(style.opacity || '1') < 0.1) continue;

                    // Tooltip's first child div text
                    var txt = '';
                    var firstDiv = tip.querySelector('div');
                    if (firstDiv) {
                        txt = (firstDiv.innerText || firstDiv.textContent || '').trim();
                    }
                    if (!txt) {
                        var fullTxt = (tip.innerText || tip.textContent || '').trim();
                        if (fullTxt) txt = fullTxt.split('\\n')[0].trim();
                    }
                    if (!txt) continue;

                    // Distance from this tooltip's center to the appointment center
                    var tcx = (r.left + r.right) / 2;
                    var tcy = (r.top + r.bottom) / 2;
                    var dist = Math.sqrt(Math.pow(tcx-apptCx,2) + Math.pow(tcy-apptCy,2));

                    if (dist < bestDist) {
                        bestDist = dist;
                        best = txt;
                    }
                }
                return best;
            """, appt_cx, appt_cy)

            log.info(f"    [Tooltip] Attempt {attempt}: '{first_line}'")

            # If still loading or empty, retry
            if not first_line:
                continue
            if re.search(r'loading', first_line, re.IGNORECASE):
                log.info("    [Tooltip] Content still loading — retrying ...")
                first_line = ""
                continue

            # Got real content
            break

        if not first_line:
            log.info("    [Tooltip] No usable tooltip after retries")
            return ""

        log.info(f"    [Tooltip] First line: '{first_line}'")

        # ── Step 3: Strip the time portion off the end ──────────────
        # Time patterns: "6:30pm", "9:30 AM", "10:00pm", etc.
        cleaned = re.sub(
            r'\s*\d{1,2}:\d{2}\s*(am|pm|AM|PM)?\s*.*$',
            '',
            first_line
        ).strip()
        cleaned = re.sub(r'[\s\-:]+$', '', cleaned).strip()

        # Safety: never return "Loading..." as an appointment type
        if cleaned and re.search(r'loading', cleaned, re.IGNORECASE):
            log.info("    [Tooltip] Cleaned value is 'Loading' — discarding")
            return ""

        if cleaned:
            log.info(f"    [Tooltip] ✓ Appointment Type: '{cleaned}'")
            return cleaned
        else:
            log.info(f"    [Tooltip] After cleaning: empty (original was '{first_line}')")
            return ""

    except Exception as e:
        log.debug(f"_get_tooltip_for_appointment error: {e}")
        return ""


def _get_therapist_for_appointment(driver, appt_element) -> str:
    """
    Returns the name of the Treating Therapist column that the given
    appointment is positioned under in the WebPT schedule.

    HTML structure in WebPT (confirmed):
      <div class="... calendar-header-slider ...">
          <span>ADELLA CHENG (5)</span>
      </div>

    Each therapist column header has the CSS class "calendar-header-slider".
    We find all of them, take the one whose X-range contains the
    appointment's X-center, and return its name (without the "(N)" count).
    """
    try:
        # ── Step 0: Scroll appointment into view ────────────────────
        try:
            driver.execute_script("""
                arguments[0].scrollIntoView({block: 'center', inline: 'center'});
            """, appt_element)
            time.sleep(0.3)
        except Exception:
            pass

        # ── Step 1: Get the appointment's bounding box ──────────────
        appt_rect = driver.execute_script("""
            var r = arguments[0].getBoundingClientRect();
            return {left: r.left, right: r.right, top: r.top, center_x: (r.left + r.right) / 2};
        """, appt_element)

        if not appt_rect:
            log.info("    [Therapist Lookup] Could not read appointment bounding box")
            return ""

        appt_left     = appt_rect.get('left', 0)
        appt_right    = appt_rect.get('right', 0)
        appt_center_x = appt_rect.get('center_x', 0)
        appt_top      = appt_rect.get('top', 0)
        log.info(f"    [Therapist Lookup] Appointment X: {appt_left:.1f}-{appt_right:.1f} (center {appt_center_x:.1f}), Top: {appt_top:.1f}")

        # ── Step 2: Find all "calendar-header-slider" elements ──────
        # This is the exact class WebPT uses for therapist column headers.
        headers = driver.execute_script("""
            var pattern = /^(.+?)\\s*\\((\\d+)\\)\\s*$/;
            var results = [];
            var nodes = document.querySelectorAll('.calendar-header-slider');
            for (var i = 0; i < nodes.length; i++) {
                var el = nodes[i];
                var txt = (el.innerText || el.textContent || '').trim();
                if (!txt) continue;
                if (txt.indexOf('\\n') >= 0) {
                    // Multi-line: take first line
                    txt = txt.split('\\n')[0].trim();
                }
                var r = el.getBoundingClientRect();
                if (r.width < 5) continue;

                var m = txt.match(pattern);
                var name = m ? m[1].trim() : txt;
                var count = m ? m[2] : '';

                results.push({
                    name: name,
                    count: count,
                    left: r.left,
                    right: r.right,
                    top: r.top,
                    width: r.width,
                    fullText: txt
                });
            }
            return results;
        """)

        if not headers:
            log.info("    [Therapist Lookup] No '.calendar-header-slider' elements found")
            return ""

        log.info(f"    [Therapist Lookup] Found {len(headers)} therapist column header(s):")
        for h in headers[:20]:
            log.info(f"      • '{h['fullText']}' [X: {h['left']:.0f}-{h['right']:.0f}, Top: {h['top']:.0f}]")

        # ── Step 3: Match by the appointment's LEFT EDGE ────────────
        # In WebPT, an appointment block ALWAYS starts at the left edge
        # of its own column. If an appointment is wide (spanning visually
        # into the next column), its therapist is still the one whose
        # column the LEFT edge falls into — NOT the one with more overlap.
        #
        # So: find the header whose X-range contains the appointment's
        # LEFT edge (with a small tolerance for sub-pixel rounding).
        TOL = 5  # pixels of tolerance

        left_matches = []
        for h in headers:
            if (h['left'] - TOL) <= appt_left <= (h['right'] + TOL):
                left_matches.append(h)

        if left_matches:
            # If the left edge lands in more than one header's range
            # (overlapping ranges at the boundary), pick the one whose
            # left edge is closest to (but not greater than) appt_left —
            # i.e. the column that actually starts at/just-before the appt.
            left_matches.sort(key=lambda h: abs(h['left'] - appt_left))
            chosen = left_matches[0]
            log.info(f"    [Therapist Lookup] ✓ MATCH (by left edge {appt_left:.0f}): '{chosen['name']}' [col X: {chosen['left']:.0f}-{chosen['right']:.0f}]")
            return chosen['name']

        # ── Step 4: Fallback — header whose left edge is closest below appt_left ──
        # Find the column whose left edge is the largest value that is
        # still <= appt_left (i.e. the column the appointment starts in).
        below = [h for h in headers if h['left'] <= appt_left + TOL]
        if below:
            below.sort(key=lambda h: h['left'])
            chosen = below[-1]  # largest left <= appt_left
            log.info(f"    [Therapist Lookup] ✓ MATCH (left edge nearest column start): '{chosen['name']}'")
            return chosen['name']

        # ── Step 5: Fallback — maximum overlap ──────────────────────
        def _overlap(a_left, a_right, b_left, b_right):
            lo = max(a_left, b_left)
            hi = min(a_right, b_right)
            return max(0.0, hi - lo)

        best_header   = None
        best_overlap  = 0.0
        for h in headers:
            ov = _overlap(appt_left, appt_right, h['left'], h['right'])
            if ov > best_overlap:
                best_overlap = ov
                best_header  = h

        if best_header and best_overlap > 0:
            log.info(f"    [Therapist Lookup] ✓ MATCH (max overlap {best_overlap:.0f}px): '{best_header['name']}'")
            return best_header['name']

        # ── Step 6: Fallback — closest header by X distance ─────────
        log.info("    [Therapist Lookup] No edge/overlap match — using closest by X")
        closest = None
        closest_dist = float('inf')
        for h in headers:
            h_center = (h['left'] + h['right']) / 2
            dist = abs(h_center - appt_center_x)
            if dist < closest_dist:
                closest_dist = dist
                closest = h

        if closest:
            log.info(f"    [Therapist Lookup] ✓ Closest: '{closest['name']}' (X dist: {closest_dist:.0f})")
            return closest['name']

    except Exception as e:
        log.debug(f"_get_therapist_for_appointment error: {e}")
        import traceback as _tb
        for line in _tb.format_exc().splitlines():
            log.debug(f"  {line}")

    return ""


def find_patient_in_schedule(driver, full_name: str, active_filter: str = ""):
    """
    Returns a list of (status, appt_date, therapist) tuples — one per appointment.

    active_filter: if a status filter ("No Show" or "Cancelled") is currently
                   active, pass it here. When the patient is found while a
                   filter is active, that filter's name BECOMES the status —
                   even if no icon is visible. This is critical for No Show
                   patients whose '?' icon is sometimes rendered as an image
                   that doesn't appear in the visible text.

    Name matching is robust to middle name differences:
      Excel:    "FAGAINS, MICHAEL KEITH"   first=MICHAEL KEITH, last=FAGAINS
      Schedule: "FAGAINS, MICHAEL K. ..."  ← only initial!
      We match using the FIRST WORD of first_name (MICHAEL) + last_name,
      and strip dots so "K." becomes "K".
    """
    first_name  = get_first_name(full_name)
    last_name   = _get_last_name(full_name)
    first_name  = re.sub(r'[,،‚]', '', first_name).strip()
    last_name   = re.sub(r'[,،‚]', '', last_name).strip()

    # Use only the FIRST WORD of the first name, in case Excel has
    # the middle name written out but the schedule shows only an initial.
    # e.g. Excel "MICHAEL KEITH" vs schedule "MICHAEL K." → match "MICHAEL"
    first_name_primary = first_name.split()[0] if first_name.split() else first_name

    log.info(f"  →  Searching schedule: First='{first_name_primary}' Last='{last_name}'  (full first name: '{first_name}')")
    time.sleep(2)

    def _norm(s: str) -> str:
        if not s:
            return ""
        s = str(s)
        s = re.sub(r'[،‚]', ',', s)
        s = re.sub(r"['`’]", '', s)
        # Strip dots (so "K." matches "K")
        s = s.replace('.', '')
        s = re.sub(r'\s+', ' ', s).strip()
        return s.upper()

    first_norm = _norm(first_name_primary)
    last_norm  = _norm(last_name)

    try:
        content_div = driver.find_element(By.ID, "contentWindow")
        page_text   = content_div.get_attribute("innerText") or ""
    except Exception:
        page_text = driver.execute_script("return document.body.innerText || '';")

    page_text_norm = _norm(page_text)

    matched_in_page = False
    if first_norm and first_norm in page_text_norm and last_norm and last_norm in page_text_norm:
        matched_in_page = True
    else:
        first_alt = first_norm.replace('-', ' ')
        last_alt  = last_norm.replace('-', ' ')
        page_alt  = page_text_norm.replace('-', ' ')
        if first_alt in page_alt and last_alt in page_alt:
            matched_in_page = True
            log.info(f"  →  Matched via hyphen-normalization")

    if not matched_in_page:
        warn(f"'{full_name}' not found in schedule page → Not Found")
        log.info(f"  →  Searched for First='{first_norm}' Last='{last_norm}'")
        return [("Not Found", "", "", "")]

    ok(f"'{full_name}' confirmed in schedule page")

    appointment_specs = []

    try:
        title_divs = driver.find_elements(By.CSS_SELECTOR, "div.title")
        log.info(f"  →  Scanning {len(title_divs)} div.title elements ...")

        for idx, el in enumerate(title_divs):
            try:
                txt = el.text.strip()
                if not txt:
                    continue
                txt_norm = _norm(txt)

                matched_here = (first_norm in txt_norm and last_norm in txt_norm)

                if not matched_here:
                    txt_alt   = txt_norm.replace('-', ' ')
                    first_alt = first_norm.replace('-', ' ')
                    last_alt  = last_norm.replace('-', ' ')
                    matched_here = (first_alt in txt_alt and last_alt in txt_alt)

                if not matched_here:
                    continue

                log.info(f"  →  Found match #{len(appointment_specs)+1}: '{txt[:80]}'")

                inner_html = ""
                try:
                    inner_html = el.get_attribute("innerHTML") or ""
                except Exception:
                    pass

                icon_status = None

                for marker in ["√", "✓", "✔", "checked in", "check-in"]:
                    if marker in txt or marker.lower() in inner_html.lower():
                        icon_status = "Checked In"
                        log.info(f"  →  Icon '{marker}' detected → Checked In")
                        break

                if not icon_status:
                    for marker in ["«", "»", "✗", "checked out", "check-out"]:
                        if marker in txt or marker.lower() in inner_html.lower():
                            icon_status = "Checked Out"
                            log.info(f"  →  Icon '{marker}' detected → Checked Out")
                            break

                if not icon_status:
                    for marker in ["Φ", "φ", "ϕ", "cancelled", "canceled"]:
                        if marker in txt or marker.lower() in inner_html.lower():
                            icon_status = "Cancelled"
                            log.info(f"  →  Icon '{marker}' detected → Cancelled")
                            break

                if not icon_status:
                    # ── No Show detection ──────────────────────────────
                    # The '?' icon appears next to the patient name, often
                    # ATTACHED to it without a space, e.g. "?Floresluna, Margarita"
                    # We need to detect '?' that is NOT inside a question
                    # (i.e. not at end of sentence). The icon is always:
                    #   - At the START of the text/name, OR
                    #   - Right after a separator like "(", ":", or whitespace
                    #
                    # Strong signals (any of these = No Show):
                    #   - txt starts with '?'
                    #   - txt contains '? ' or ' ?'
                    #   - inner_html contains '?' near the start
                    #   - "no show" / "no-show" anywhere in text
                    is_no_show = False
                    txt_stripped = txt.strip()

                    if txt_stripped.startswith('?'):
                        is_no_show = True
                        log.info(f"  →  '?' at start of text → No Show")
                    elif re.search(r'(^|[\s\(])\?', txt):
                        is_no_show = True
                        log.info(f"  →  '?' icon detected → No Show")
                    elif "no show" in txt.lower() or "no-show" in txt.lower():
                        is_no_show = True
                        log.info(f"  →  'no show' text → No Show")
                    else:
                        # Check inner HTML for an early '?' that's not part
                        # of a longer string (DOB, ID, etc.)
                        # Look at first 200 chars of inner_html for ?
                        early_html = inner_html[:200] if inner_html else ""
                        if '?' in early_html:
                            # Make sure it's not just a URL parameter or attribute
                            # Look for ? as actual content (not href="?...")
                            # Strategy: find '?' positions and check if they're
                            # inside text content (between > and <)
                            content_q = re.search(r'>([^<]*?\?[^<]*?)<', early_html)
                            if content_q:
                                is_no_show = True
                                log.info(f"  →  '?' in inner HTML content → No Show")

                    if is_no_show:
                        icon_status = "No Show"

                # ── ACTIVE FILTER FALLBACK ─────────────────────────────
                # If we're currently viewing a "No Show" or "Cancelled"
                # filter and the patient appears in that filtered list,
                # then their status MUST be that filter's name. The icon
                # may not be visible (sometimes WebPT renders the '?' or
                # 'Φ' as an image that doesn't appear in the text).
                # The fact that they're listed under this filter is itself
                # the strongest evidence of their status.
                if not icon_status and active_filter:
                    if active_filter.lower() in ("no show", "no-show", "noshow"):
                        icon_status = "No Show"
                        log.info(f"  →  Active filter is 'No Show' → Status: No Show (icon not visible)")
                    elif active_filter.lower() in ("cancelled", "canceled"):
                        icon_status = "Cancelled"
                        log.info(f"  →  Active filter is 'Cancelled' → Status: Cancelled (icon not visible)")

                if not icon_status:
                    icon_status = "Other"
                    log.info(f"  →  No icon detected → Status will be 'Other'")

                appointment_specs.append({
                    "icon_status": icon_status,
                    "title_text":  txt,
                    "dom_index":   idx,
                })
            except Exception:
                continue
    except Exception as e:
        log.debug(f"div.title scan error: {e}")

    if not appointment_specs:
        warn(f"Could not find appointment div for '{full_name}' → Not Found")
        return [("Not Found", "", "", "")]

    log.info(f"  →  Total appointments to process for '{full_name}': {len(appointment_specs)}")

    results = []
    for spec_idx, spec in enumerate(appointment_specs):
        log.info(f"  ━━━ Processing appointment {spec_idx+1}/{len(appointment_specs)} ━━━")

        target_text = spec["title_text"]
        matched_el = None
        try:
            current_titles = driver.find_elements(By.CSS_SELECTOR, "div.title")
            matches_so_far = 0
            for el in current_titles:
                try:
                    if el.text.strip() == target_text:
                        if matches_so_far == spec_idx or len(appointment_specs) == 1:
                            matched_el = el
                            break
                        matches_so_far += 1
                except Exception:
                    continue
        except Exception as e:
            log.debug(f"Re-finding element error: {e}")

        if matched_el is None:
            warn(f"Could not re-find appointment {spec_idx+1} — skipping")
            continue

        result = _process_single_appointment(
            driver, full_name, matched_el, spec["icon_status"]
        )
        results.append(result)

        _ensure_on_schedule_view(driver)
        time.sleep(1)

    _ensure_on_schedule_view(driver)

    if not results:
        return [("Not Found", "", "", "")]

    return results


def _click_day_view(driver):
    log.info("  →  Clicking 'Day' view button ...")
    for xpath in [
        "//input[@value='Day']",
        "//button[normalize-space(text())='Day']",
        "//a[normalize-space(text())='Day']",
        "//td[normalize-space(text())='Day']",
        "//*[normalize-space(text())='Day']",
    ]:
        try:
            els = driver.find_elements(By.XPATH, xpath)
            for el in els:
                if el.is_displayed():
                    driver.execute_script("arguments[0].click();", el)
                    time.sleep(2)
                    ok("'Day' view selected")
                    return
        except Exception:
            continue
    warn("'Day' button not found — continuing with current view")


def _is_filter_active(driver, status: str) -> bool:
    """
    Returns True if the given filter button ('Checked Out', 'No Show',
    'Cancelled') currently appears to be active/pressed.

    WebPT styles active filters differently (usually pressed state, or
    'x-btn-pressed' class, or different background color). We check
    common Ext JS patterns.
    """
    try:
        # Find the button/element with this filter's text
        for xpath in [
            f"//input[@value='{status}']",
            f"//button[normalize-space(text())='{status}']",
            f"//a[normalize-space(text())='{status}']",
            f"//td[normalize-space(text())='{status}']",
            f"//*[normalize-space(text())='{status}']",
        ]:
            try:
                els = driver.find_elements(By.XPATH, xpath)
                for el in els:
                    if not el.is_displayed():
                        continue

                    # Check the element itself and its ancestors for
                    # pressed/active class indicators.
                    for target in [el] + el.find_elements(By.XPATH, "./ancestor::*[position()<=3]"):
                        try:
                            cls = (target.get_attribute("class") or "").lower()
                            if any(k in cls for k in [
                                "pressed", "active", "selected",
                                "x-btn-pressed", "x-pressed"
                            ]):
                                log.info(f"  →  Filter '{status}' is ACTIVE (class='{cls[:60]}')")
                                return True
                        except Exception:
                            continue

                    # If we found the button but no active class → not active
                    return False
            except Exception:
                continue
    except Exception:
        pass
    # Default: assume not active (safer to click and toggle if needed)
    return False


def _click_status_filter(driver, status: str):
    log.info(f"  →  Clicking '{status}' filter button ...")
    for xpath in [
        f"//input[@value='{status}']",
        f"//button[normalize-space(text())='{status}']",
        f"//a[normalize-space(text())='{status}']",
        f"//td[normalize-space(text())='{status}']",
        f"//*[normalize-space(text())='{status}']",
    ]:
        try:
            els = driver.find_elements(By.XPATH, xpath)
            for el in els:
                if el.is_displayed():
                    driver.execute_script("arguments[0].click();", el)
                    time.sleep(2)
                    ok(f"'{status}' filter selected")
                    return
        except Exception:
            continue
    warn(f"'{status}' button not found — continuing with current view")


# =====================================================================
# FULL SCHEDULE CHECK FLOW
# =====================================================================

def _is_on_schedule_view(driver) -> bool:
    """
    Returns True if the browser is currently on the View Schedule page.
    Detection is based on the presence of the ClinicChange dropdown element,
    which only exists on the schedule page.
    """
    try:
        el = driver.find_element(By.ID, "ClinicChange")
        return el is not None
    except Exception:
        return False


def check_schedule_status(driver, wait, full_name: str, target_date: str, clinic_name: str):
    """
    Returns a list of (status, appt_date) tuples — one per appointment.

    Flow per patient:
      1. Verify we're on View Schedule (open it once if not)
      2. Select Clinic (skipped if same as last)
      3. Navigate to Date in the calendar
      4. Click Day view
      5. Search for the patient
      6. Process found appointments (close popup at end)

    The script stays on View Schedule for the whole run — it does NOT
    return to the dashboard between patients.
    """
    step(f"Schedule Check | '{full_name}' | {target_date} | {clinic_name}")

    # ── Step 1: Make sure we're on View Schedule ──────────────────
    if _is_on_schedule_view(driver):
        ok("Already on View Schedule — staying here")
    else:
        warn("Not on View Schedule — re-opening it")
        opened = _open_view_schedule(driver, wait)
        if not opened:
            warn("Could not open View Schedule — returning Other")
            return [("Other", "", "", "")]

    # ── Step 2: Select clinic (skipped internally if same as last) ──
    select_clinic_in_schedule(driver, clinic_name)

    # ── Step 3: Navigate to the target date in the calendar ─────────
    navigate_schedule_to_date(driver, target_date)

    # ── Step 4: Make sure we're in Day view ─────────────────────────
    _click_day_view(driver)

    # ── Verify date actually changed ────────────────────────────────
    try:
        time.sleep(1)
        page_body = driver.find_element(By.TAG_NAME, "body").text[:2000]
        try:
            target_d = pd.to_datetime(target_date)
            target_day_str = str(target_d.day)
            target_month_abbr = target_d.strftime("%b")
            target_year_str = str(target_d.year)
            shows_target = (target_month_abbr in page_body and target_day_str in page_body and target_year_str in page_body)
            if not shows_target:
                log.info(f"  →  Day view doesn't seem to show target date — navigating again")
                navigate_schedule_to_date(driver, target_date)
                _click_day_view(driver)
                time.sleep(1)
        except Exception as e:
            log.debug(f"  →  Date verification skipped: {e}")
    except Exception as e:
        log.debug(f"  →  Day view verification error: {e}")

    # ── Step 5: Activate filters (only if not already active) ───────
    # Per user request: keep No Show + Cancelled active across patients
    # in the SAME clinic. When the clinic changes, _extra_filters_active
    # is reset by select_clinic_in_schedule, so we re-activate them here.
    global _extra_filters_active

    if not _extra_filters_active:
        log.info("  →  Activating filters for the first time in this clinic ...")
        if not _is_filter_active(driver, "Checked Out"):
            log.info("  →  Activating 'Checked Out' filter")
            _click_status_filter(driver, "Checked Out")
            time.sleep(0.8)

        log.info("  →  Activating 'No Show' filter")
        _click_status_filter(driver, "No Show")
        time.sleep(0.8)

        log.info("  →  Activating 'Cancelled' filter")
        _click_status_filter(driver, "Cancelled")
        time.sleep(1.2)

        _extra_filters_active = True
    else:
        log.info("  →  Filters (Checked Out + No Show + Cancelled) already active — reusing")

    # ── Step 6: Search for the patient (all 3 statuses visible) ─────
    # The icon next to the name determines the status.
    results = find_patient_in_schedule(driver, full_name, active_filter="")

    # ── Step 7: Keep filters active for the next patient ────────────
    # We do NOT toggle off No Show / Cancelled. They stay on so the next
    # patient (same clinic) can be searched immediately. They'll be reset
    # automatically when the clinic changes.

    return results


# =====================================================================
# PROCESS ONE PATIENT
# =====================================================================

def process_patient(driver, wait, patient_id: str, clinic_name: str,
                    target_date: str, full_name: str) -> list:
    step(f"Processing | ID:{patient_id} | {clinic_name} | {target_date} | {full_name}")

    appointments = check_schedule_status(driver, wait, full_name, target_date, clinic_name)

    out_records = []
    for idx, appt in enumerate(appointments):
        # Each appt is (status, appt_date, therapist, appt_type)
        status, appt_date, therapist, appt_type = appt
        if len(appointments) > 1:
            log.info(f"  →  Appointment {idx+1}/{len(appointments)} → Status: {status}")

        ok(f"STATUS for '{full_name}' (appt {idx+1}): [{status}] | Date: [{appt_date}] | Therapist: [{therapist}] | Type: [{appt_type}]")

        out_records.append({
            "EMR_ID":    patient_id,
            "Name":      full_name,
            "Clinic":    clinic_name,
            "Date":      target_date,
            "Status":    status,
            "Appt_Date": appt_date,
            "Therapist": therapist,
            "ApptType":  appt_type,
        })

    return out_records


# =====================================================================
# SAVE TO EXCEL — 7 COLUMNS
# =====================================================================

def save_styled_excel(results: list, output_path: str):
    """
    Header columns (9):
      A  Patient ID
      B  Patient Name
      C  Clinic Name
      D  Appointment Date
      E  Status
      F  Status Update Date
      G  Status Update Time
      H  Treating Therapist
      I  Appointment Type
    """
    headers = [
        "Patient ID",
        "Patient Name",
        "Clinic Name",
        "Appointment Date",
        "Status",
        "Status Update Date",
        "Status Update Time",
        "Treating Therapist",
        "Appointment Type",
    ]

    black_fill   = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    white_bold   = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_side    = Side(style="thin", color="000000")
    thin_border  = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    data_font    = Font(name="Calibri", size=11)

    file_exists = os.path.exists(output_path)
    if file_exists:
        # ── Delete old file and start fresh ────────────────────────
        # User wants the results file to ALWAYS contain only the
        # current run's data — no old rows kept.
        try:
            os.remove(output_path)
            log.info(f"  →  Deleted old file: {output_path}")
        except PermissionError:
            warn(f"Could not delete old file (open in Excel?): {output_path}")
            # Will save with timestamp suffix at the end
        except Exception as e:
            warn(f"Could not delete old file: {e}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    next_row = 1
    need_headers = True
    log.info("  →  Creating fresh output file")

    if need_headers:
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.fill      = black_fill
            cell.font      = white_bold
            cell.alignment = center_align
            cell.border    = thin_border
        next_row = 2

    rows_added = 0
    for record in results:
        appt_full = record.get("Appt_Date", "")
        if appt_full and " " in appt_full:
            parts = appt_full.split(" ", 1)
            try:
                appt_date_part = pd.to_datetime(parts[0]).strftime("%m/%d/%Y")
            except Exception:
                appt_date_part = parts[0]
            appt_time_part = parts[1] if len(parts) > 1 else ""
        elif appt_full:
            try:
                appt_date_part = pd.to_datetime(appt_full).strftime("%m/%d/%Y")
            except Exception:
                appt_date_part = appt_full
            appt_time_part = ""
        else:
            appt_date_part = ""
            appt_time_part = ""

        raw_date = record.get("Date", "")
        try:
            appt_date_fmt = pd.to_datetime(raw_date).strftime("%m/%d/%Y")
        except Exception:
            appt_date_fmt = raw_date

        emr_id_raw = record.get("EMR_ID", "")
        try:
            emr_id_val = int(emr_id_raw) if str(emr_id_raw).strip() else ""
        except (ValueError, TypeError):
            emr_id_val = emr_id_raw

        values = [
            emr_id_val,                       # A: Patient ID
            record.get("Name",      ""),      # B: Patient Name
            record.get("Clinic",    ""),      # C: Clinic Name
            appt_date_fmt,                    # D: Appointment Date
            record.get("Status",    ""),      # E: Status
            appt_date_part,                   # F: Status Update Date
            appt_time_part,                   # G: Status Update Time
            record.get("Therapist", ""),      # H: Treating Therapist
            record.get("ApptType",  ""),      # I: Appointment Type
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=next_row, column=col_idx, value=value)
            cell.font      = data_font
            cell.alignment = center_align
            cell.border    = thin_border
        next_row += 1
        rows_added += 1

    if need_headers:
        col_widths = {
            "A": 12, "B": 25, "C": 18, "D": 16,
            "E": 15, "F": 18, "G": 18, "H": 22, "I": 20,
        }
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

    for row_num in range(next_row - rows_added, next_row):
        ws.row_dimensions[row_num].height = 18

    try:
        wb.save(output_path)
        ok(f"Saved {rows_added} row(s) to: {output_path}")
    except PermissionError:
        ts_path = output_path.replace('.xlsx', f'_{int(time.time())}.xlsx')
        wb.save(ts_path)
        warn(f"File was locked. Saved as: {ts_path}")


# =====================================================================
# MAIN
# =====================================================================

ALLOWED_CLINICS = [
    "Allerton", "Belmont", "Bushwick", "Astoria", "Bay Ridge", "Bedstuy",
    "BensonHurst - Adults", "Central Harlem", "Chelsea", "Midtown", "Castle Hill",
    "Church Ave", "Clinton Hill", "Cobble Hill", "Lenox hill", "Upper East Side",
    "Grand Concourse", "Greenpoint", "Tribeca", "Brownsville", "Corona", "Flatbush",
    "Hell's Kitchen", "Huntspoint", "Midwood", "Jamaica", "Sugar Hill",
    "Riverdale", "Sunset - Adults", "Fordham", "Inwood", "Sheepshead Bay",
] 
CLINIC_COUNTS_OUTPUT_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "clinic_pt_counts.xlsx"
)


def save_clinic_counts_excel(records):
    details = pd.DataFrame(records)
    successful = details[details["Status"] == "OK"].copy()
    successful["Count"] = pd.to_numeric(successful["Count"], errors="raise")
    # Keep failed clinics visible with an empty total, rather than reporting zero.
    totals = details.groupby("Clinic", sort=False).agg(
        Status=("Status", "first"), Entries=("Count", "count")
    )
    totals["Total"] = successful.groupby("Clinic", sort=False)["Count"].sum()
    # Clinic is the selected parent clinic; Schedule Column identifies its services.
    # Thus all Allerton service columns contribute to the single Allerton total.
    totals["Included Schedule Columns"] = successful.groupby("Clinic", sort=False)[
        "Schedule Column"
    ].agg(lambda columns: "; ".join(dict.fromkeys(
        str(column).strip() for column in columns if pd.notna(column) and str(column).strip()
    )))
    by_column = successful.groupby(
        ["Clinic", "Schedule Column"], sort=False, as_index=False, dropna=False
    )["Count"].sum().rename(columns={"Count": "Total"})
    by_type = successful.groupby(["Clinic", "Count Type"], sort=False, as_index=False)["Count"].sum()
    by_type = by_type.rename(columns={"Count": "Total"})
    with pd.ExcelWriter(CLINIC_COUNTS_OUTPUT_PATH, engine="openpyxl") as writer:
        details.to_excel(writer, sheet_name="Details", index=False)
        totals.reset_index().to_excel(writer, sheet_name="Clinic Totals", index=False)
        by_column.to_excel(writer, sheet_name="Totals by Schedule Column", index=False)
        by_type.to_excel(writer, sheet_name="Totals by Type", index=False)


def scrape_clinic_pt_counts(driver, wait):
    global _current_clinic

    def read_pt_label(d):
        return d.execute_script(r"""
            const visible = el => el.getClientRects().length &&
                getComputedStyle(el).visibility !== 'hidden';
            // Accept every header category, rather than a fixed list of names.
            const countPattern = /^([^()\r\n]+?)\s*\(\s*(\d+)\s*\)$/;
            const labelsIn = el => (el.innerText || '').split(/\r?\n/)
                .map(text => text.trim()).filter(text => countPattern.test(text));
            // Count labels can be rendered in divs, links or table cells as well as spans.
            // Keep the smallest element for each label to avoid counting ancestors twice.
            const countElements = Array.from(document.querySelectorAll('body *'))
                .filter(el => visible(el) && !['SCRIPT', 'STYLE', 'OPTION'].includes(el.tagName))
                .flatMap(el => labelsIn(el).filter(label =>
                    !Array.from(el.children).some(child => visible(child) &&
                        labelsIn(child).includes(label)))
                    .map(label => ({el, label})));
            const datePattern = /\b(?:Mon(?:day)?|Tue(?:sday)?|Wed(?:nesday)?|Thu(?:rsday)?|Fri(?:day)?|Sat(?:urday)?|Sun(?:day)?)\s+\d{1,2}\/\d{1,2}(?:\/\d{2,4})?\b/i;
            const dates = Array.from(document.querySelectorAll('div,span,th,td'))
                .filter(el => visible(el) && datePattern.test(el.innerText || '') &&
                    !Array.from(el.children).some(child => datePattern.test(child.innerText || '')))
                .map(el => {
                    // A centered heading may be narrower than its two schedule columns.
                    // Expand to its enclosing day container, stopping before another date.
                    let container = el;
                    for (let parent = el.parentElement; parent; parent = parent.parentElement) {
                        const headings = (parent.innerText || '').match(
                            new RegExp(datePattern.source, 'gi')) || [];
                        if (headings.length !== 1) break;
                        container = parent;
                    }
                    return {text: el.innerText.match(datePattern)[0],
                            rect: container.getBoundingClientRect(),
                            headingRect: el.getBoundingClientRect()};
                });
            return countElements.map(({el, label}) => {
                    const rect = el.getBoundingClientRect();
                    let title = '', header = label;
                    // Preserve the full surrounding column header and all its categories.
                    for (let parent = el.parentElement, depth = 0;
                         parent && depth < 6; parent = parent.parentElement, depth++) {
                        const text = (parent.innerText || '').trim();
                        let other = text;
                        for (const item of countElements) {
                            if (parent.contains(item.el)) other = other.replace(item.label, '');
                        }
                        other = other.trim();
                        const titles = other.split(/\n/).map(s => s.trim()).filter(Boolean);
                        if (titles.length > 1) break;
                        if (other && !datePattern.test(other)) {
                            title = other; header = text; break;
                        }
                    }
                    const center = rect.left + rect.width / 2;
                    const candidates = dates.filter(date => {
                        const r = date.rect;
                        return date.headingRect.top <= rect.top &&
                            r.left <= center && r.right >= center;
                    }).sort((a, b) => (a.rect.width - b.rect.width) ||
                        b.headingRect.top - a.headingRect.top);
                    const date = candidates.length ? candidates[0].text : '';
                    const match = label.match(countPattern);
                    return {title, header, date, label, type: match[1].trim(),
                            count: Number(match[2]), x: rect.left, y: rect.top};
                }).sort((a, b) => a.y - b.y || a.x - b.x)
                .map(({x, y, ...entry}) => entry);
        """) or False

    records = []
    for clinic_name in ALLOWED_CLINICS:
        step(f"Reading PT count | {clinic_name}")
        record = {"Clinic": clinic_name, "Entry": None, "Schedule Day": "", "Schedule Date": "",
                  "Schedule Column": "", "Header Text": "", "Count Type": "",
                  "Count Label": "", "Count": None,
                  "Captured At": "", "Status": ""}
        try:
            select_element = _get_clinic_select_element(driver)
            if select_element is None:
                raise RuntimeError("Scheduling clinic dropdown was not found.")
            options = driver.execute_script(
                "return Array.from(arguments[0].options)"
                ".filter(o => !o.disabled && o.value)"
                ".map(o => ({name: o.text.trim(), value: o.value}));",
                select_element,
            )
            matches = [o for o in options if o["name"].casefold() == clinic_name.casefold()]
            if len(matches) != 1:
                raise RuntimeError(f"Clinic '{clinic_name}' was not uniquely found in the dropdown.")
            target = matches[0]
            # Recheck the actual dropdown rather than relying on the cached clinic.
            _current_clinic = None
            select_clinic_in_schedule(driver, target["name"])

            def selected_clinic(d):
                element = _get_clinic_select_element(d)
                return element is not None and d.execute_script(
                    "return arguments[0].value;", element
                ) == target["value"]

            wait.until(selected_clinic)
            wait_for_page_stable(driver, timeout=WAIT_TIMEOUT)
            last_label = None
            stable_since = None

            def loaded_pt_label(d):
                nonlocal last_label, stable_since
                busy = d.execute_script("""
                    var extBusy = window.Ext && Ext.Ajax &&
                        typeof Ext.Ajax.isLoading === 'function' && Ext.Ajax.isLoading();
                    var masks = document.querySelectorAll(
                        '.x-mask-loading, .x-loading-mask, .x-loadmask');
                    return Boolean(extBusy) || Array.from(masks).some(function(el) {
                        return el.getClientRects().length &&
                            getComputedStyle(el).visibility !== 'hidden';
                    });
                """)
                if busy or not selected_clinic(d):
                    last_label = None
                    stable_since = None
                    return False
                result = read_pt_label(d)
                if not result:
                    last_label = None
                    stable_since = None
                    return False
                if result != last_label:
                    last_label = result
                    stable_since = time.monotonic()
                    return False
                return result if time.monotonic() - stable_since >= 1.0 else False

            entries = wait.until(loaded_pt_label)
            # Some schedule views render columns only when scrolled into view.
            # Sweep horizontal panes and merge the same day/column/category once.
            collected = {}

            def collect_entries(items):
                for item in items:
                    key = (item["date"], item["title"], item["type"], item["label"])
                    collected.setdefault(key, item)

            collect_entries(entries)
            scroll_panes = driver.execute_script("""
                return Array.from(document.querySelectorAll('div')).filter(el => {
                    var style = getComputedStyle(el);
                    return el.clientWidth > 0 && el.scrollWidth > el.clientWidth + 20 &&
                        /auto|scroll/.test(style.overflowX) && el.getClientRects().length;
                });
            """)
            for pane in scroll_panes:
                try:
                    original, maximum, width = driver.execute_script(
                        "return [arguments[0].scrollLeft, "
                        "arguments[0].scrollWidth - arguments[0].clientWidth, "
                        "arguments[0].clientWidth];", pane
                    )
                    try:
                        position = 0
                        while True:
                            driver.execute_script("arguments[0].scrollLeft = arguments[1];", pane, position)
                            # Allow scroll-triggered rendering before polling for stable counts.
                            time.sleep(0.5)
                            collect_entries(wait.until(loaded_pt_label))
                            if position >= maximum:
                                break
                            position = min(maximum, position + max(1, int(width * 0.8)))
                    finally:
                        driver.execute_script("arguments[0].scrollLeft = arguments[1];", pane, original)
                except StaleElementReferenceException:
                    warn("Schedule pane was replaced while scrolling; retained captured entries")
            entries = list(collected.values())
            captured_at = pd.Timestamp.now().isoformat()
            for index, entry in enumerate(entries, start=1):
                records.append({**record, "Entry": index,
                                "Schedule Day": entry["date"],
                                "Schedule Date": entry["date"].split()[-1] if entry["date"] else "",
                                "Schedule Column": entry["title"],
                                "Header Text": entry["header"],
                                "Count Type": entry["type"],
                                "Count Label": entry["label"], "Count": entry["count"],
                                "Captured At": captured_at, "Status": "OK"})
            ok(f"{clinic_name}: captured {len(entries)} schedule header entries")
        except Exception as exc:
            message = str(exc).splitlines()
            record["Status"] = f"Error: {type(exc).__name__}: {message[0] if message else 'Unknown error'}"
            fail(f"Could not read PT count for '{clinic_name}': {record['Status']}")
            records.append(record)
        save_clinic_counts_excel(records)
    ok(f"Clinic PT counts saved to: {CLINIC_COUNTS_OUTPUT_PATH}")


def main():
    step("WEBPT CLINIC PT COUNTS - STARTING")
    driver = create_driver()
    wait = WebDriverWait(driver, WAIT_TIMEOUT)
    try:
        login(driver, wait)
        scrape_clinic_pt_counts(driver, wait)
    finally:
        step("Closing browser")
        driver.quit()
        ok("Done")


if __name__ == "__main__":
    main()
