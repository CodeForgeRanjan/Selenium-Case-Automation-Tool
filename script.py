from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os
from datetime import datetime

# CONFIGURATION FOR TOOL 

FILE_NAME = os.getenv("FILE_NAME", "Data_VS_Code.xlsx")
COLUMN_NAME = "caseNo"
FAILED_FILE = "failed_cases.xlsx"
LOG_FILE = "automation_log.txt"

# URL

BASE_URL = ""
HOME_URL = ""

USER_DATA_DIR = os.path.join(os.getcwd(), "chrome_profile")

# Enable/Disable profile
USE_PROFILE = False  

WAIT_TIME = 15
RESTART_EVERY = 100


def log(msg):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(msg)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {msg}\n")

def start_driver():
    options = webdriver.ChromeOptions()
    
    
# Use profile only if enabled
    if USE_PROFILE:
        options.add_argument(f"--user-data-dir={USER_DATA_DIR}")
        options.add_argument("--profile-directory=Default")

    
    # hide for browser log
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        return driver
    except Exception as e:
        log(f"Driver start error: {e}")
        raise e

#  START 
log("SCRIPT STARTED")

#  Start Driver 
driver = start_driver()

#Open websites
try:
    driver.get(HOME_URL)
    log("Browser opened successfully.")
except Exception as e:
    log(f"URL error: {e}")

# Manual Intervention for Login
input("\n Login and press enter")

# Load Excel
try:
    df = pd.read_excel(FILE_NAME)
    log(f"Excel loaded: {len(df)} cases found.")
except Exception as e:
    log(f"Excel error: {e}")
    exit()

failed_cases = []

for index, caseNo in enumerate(df[COLUMN_NAME], start=1):
    try:
        caseNo = str(caseNo).strip()
        if not caseNo or caseNo == "nan":
            continue
            
        log(f"[{index}/{len(df)}] Processing: {caseNo}")

        # Navigate to Case URL
        driver.get(BASE_URL.format(caseNo))
        wait = WebDriverWait(driver, WAIT_TIME)

        #Dropdown selection
        dropdown = wait.until(
            EC.presence_of_element_located((By.ID, "CaseCOMMENTCATEGORY"))
        )
        Select(dropdown).select_by_visible_text("Bridge BUG")

        #Comment typing
        comment_box = wait.until(
            EC.presence_of_element_located((By.ID, "case_comment"))
        )
        comment_box.clear()
        comment_box.send_keys("Blank checks raised")

        #Submit Click
        submit_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='submit']"))
        )
        submit_btn.click()

        time.sleep(1.5) 
        log(f"Case {caseNo} Completed.")

        # Stability restart logic
        if index % RESTART_EVERY == 0:
            log("Restarting browser to clear memory..")
            driver.quit()
            time.sleep(2)
            driver = start_driver()

    except Exception as e:
        log(f"Error in {caseNo}: {str(e)[:50]}...")
        failed_cases.append(caseNo)
        
        # If show Error then start fresh browser
        try:
            driver.quit()
        except:
            pass
        time.sleep(2)
        driver = start_driver()
        continue

# FINISH 
if failed_cases:
    pd.DataFrame(failed_cases, columns=["FAILED_CASE_ID"]).to_excel(FAILED_FILE, index=False)
    log(f"{len(failed_cases)} cases failed. Saved to {FAILED_FILE}")

driver.quit()
log("SCRIPT FINISHED SUCCESSFULLY")
