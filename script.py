import re
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from openpyxl import load_workbook
from webdriver_manager.chrome import ChromeDriverManager

ua = UserAgent()
USER_AGENT = ua.random
ChromeDriverPath = "C:/chromedriver/chromedriver.exe"

BASE_URL = 'https://www.sportsbet.com.au'
FILE_NAME = 'Race Meetings.xlsm'
target_column = 23
ALLOWED_MEETINGS = ['(VIC)', '(NSW)', '(QLD)', '(SA)', '(WA)', '(NT)', '(TAS)', '(ACT)', '(NZ)', '(NZL)']
FS = {}
SR = {}

def setup_driver():
    options = Options()
    options.headless = True
    options.add_argument("--disable-images")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-popup-blocking")
    options.add_argument(f"--user-agent={USER_AGENT}")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--no-first-run")
    options.add_argument("--disable-site-isolation-trials")

    service = Service(ChromeDriverPath)
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(800)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    return driver

def extract_sb_rating(driver, race_url):
    global SR

    driver.get(BASE_URL + race_url)
    wait = WebDriverWait(driver, 15)

    wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div[data-automation-id^='racecard-outcome-']")
        )
    )

    runner_ids = [
        el.get_attribute("data-automation-id").replace("racecard-outcome-", "")
        for el in driver.find_elements(
            By.CSS_SELECTOR,
            "div[data-automation-id^='racecard-outcome-']")
    ]

    print("--Next Race--")

    for runner_id in runner_ids:
        try:
            r = driver.find_element(
                By.CSS_SELECTOR,
                f"div[data-automation-id='racecard-outcome-{runner_id}']"
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", r
            )
            driver.execute_script("arguments[0].click();", r)

            soup = BeautifulSoup(driver.page_source, "html.parser")

            horse_el = soup.select_one(
                f"div[data-automation-id='racecard-outcome-{runner_id}'] "
                "div[data-automation-id='racecard-outcome-name'] span"
            )
            if not horse_el:
                continue

            raw_name = horse_el.get_text(strip=True)
            horse_name = re.sub(r"^\d+\.\s*", "", raw_name)

            sb_el = soup.select_one(
                f"div[data-automation-id='shortform-{runner_id}'] "
                "div[data-automation-id='shortform-SB Rating'] span:last-child"
            )
            if not sb_el:
                continue

            sb_rating = sb_el.get_text(strip=True)

            SR.setdefault("RACE", {})
            SR["RACE"][horse_name] = sb_rating

            print(f"✅ {horse_name} → SB Rating {sb_rating}")

        except Exception:
            # intentionally silent – DOM instability
            continue

def disable_international_filter(driver):
    wait = WebDriverWait(driver, 20)

    try:
        # Detect ON state
        on_state = driver.find_elements(
            By.CSS_SELECTOR,
            "div[data-automation-id='filter-button-international-on']"
        )

        if not on_state:
            print("Int'l already OFF")
            return

        print("Int'l filter ON → disabling")

        # Click the LABEL (this is critical)
        label = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "label[for='ALL_RACING_PAGEINTERNATIONAL']")
            )
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", label
        )
        time.sleep(0.3)

        driver.execute_script("arguments[0].click();", label)

        # WAIT FOR OFF STATE
        wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "div[data-automation-id='filter-button-international-off']")
            )
        )

        # WAIT FOR TABLE RE-DRAW (critical)
        wait.until(
            EC.staleness_of(
                driver.find_element(
                    By.CSS_SELECTOR,
                    "td[data-automation-id^='horse-racing-section-row-']"
                )
            )
        )

        print("Int'l filter OFF")

    except Exception as e:
        print("Failed to disable Int'l filter:", e)


def get_races(driver):
    driver.get(BASE_URL + "/racing-schedule")

    wait = WebDriverWait(driver, 20)

    # Initial load
    wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "td[data-automation-id^='horse-racing-section-row-']")
        )
    )

    # Disable Int'l properly
    disable_international_filter(driver)

    # Re-wait after filter
    wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "td[data-automation-id^='horse-racing-section-row-']")
        )
    )

    soup = BeautifulSoup(driver.page_source, "html.parser")

    cells = soup.find_all(
        "td",
        attrs={
            "data-automation-id": re.compile(
                r"^horse-racing-section-row-\d+-col-\d+-event-cell$"
            )
        },
    )

    race_links = [
        a["href"]
        for td in cells
        if (a := td.find("a", href=True))
    ]

    print(f"Filtered races (AU/NZ only): {len(race_links)}")

    return race_links

def normalize_horse(name: str) -> str:
    return (
        name.strip()
        .upper()          # case-insensitive
        .replace(".", "") # remove dots like "5. "
    )


def save_sb_to_excel(excel_file, SR):
    workbook = load_workbook(filename=excel_file, keep_vba=True)

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        for row in sheet.iter_rows(min_row=1):
            horse_cell = row[3]  # Column D
            if not horse_cell.value:
                continue

            excel_horse = normalize_horse(str(horse_cell.value))

            for sb_horse, sb_rating in SR.get("RACE", {}).items():
                if normalize_horse(sb_horse) == excel_horse:
                    sheet.cell(row=horse_cell.row, column=25, value=sb_rating)
                    print(
                        f"Saved | {horse_cell.value} → {sb_rating}"
                    )
                    break

    workbook.save(excel_file)


def main():
    driver = setup_driver()

    race_links = get_races(driver)

    for race_link in race_links:
         extract_sb_rating(driver, race_link)

    driver.quit()
    save_sb_to_excel(FILE_NAME, SR)

if __name__ == '__main__':
    main()
