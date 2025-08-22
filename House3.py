import os
import time
import shutil
import subprocess
import logging
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from dotenv import load_dotenv
import requests

# 日誌設置，同時輸出到控制台和檔案
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(module)s] %(message)s',
    handlers=[
        logging.FileHandler('house3.log'),
        logging.StreamHandler()
    ]
)

# 定義單一的下載目錄和常量
download_dir = os.path.abspath("downloads")
MAX_RETRIES = 3
DOWNLOAD_TIMEOUT = 15  # 縮短下載超時時間
WAIT_TIMEOUT = 8  # 縮短 WebDriverWait 時間

def clear_download_dirs():
    if os.path.exists(download_dir):
        shutil.rmtree(download_dir)
    os.makedirs(download_dir)
    logging.info(f"創建下載目錄: {download_dir}")

def setup_environment():
    try:
        result = subprocess.run(['chromium-browser', '--version'], capture_output=True, text=True)
        if result.returncode == 0:
            logging.info(f"Chromium 版本: {result.stdout.strip()}")
        else:
            logging.warning("無法獲取 Chromium 版本，將嘗試安裝")
            subprocess.run(['sudo', 'apt-get', 'update', '-qq'], check=True)
            subprocess.run(['sudo', 'apt-get', 'install', '-y', 'chromium-browser', 'chromium-chromedriver'], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            logging.info("Chromium 及 ChromeDriver 已安裝")
        result = subprocess.run(['pip', 'show', 'selenium'], capture_output=True, text=True)
        if "selenium" not in result.stdout or "webdriver-manager" not in subprocess.run(['pip', 'show', 'webdriver-manager'], capture_output=True, text=True).stdout:
            subprocess.run(['pip', 'install', 'selenium==4.8.0', 'webdriver-manager==3.8.5'], check=True)
            logging.info("Selenium 及 WebDriver Manager 已安裝")
        else:
            logging.info("Selenium 及 WebDriver Manager 已存在，跳過安裝")
    except subprocess.CalledProcessError as e:
        logging.error(f"環境準備失敗: {e}")
        raise

def get_chrome_options(download_dir):
    chrome_options = Options()
    chrome_options.add_argument('--headless=new')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--disable-popup-blocking')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--no-first-run')
    chrome_options.add_argument('--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_argument('--window-size=1920,1080')
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": False
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.binary_location = '/usr/bin/chromium-browser'
    return chrome_options

def wait_for_new_file(download_dir, initial_files, timeout=DOWNLOAD_TIMEOUT, extensions=('.csv', '.xlsx')):
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_files = set(f for f in os.listdir(download_dir) if f.endswith(extensions))
        new_files = current_files - initial_files
        if new_files:
            return new_files
        time.sleep(0.5)
    return set()

def wait_for_page_load(driver, timeout=WAIT_TIMEOUT):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script('return document.readyState') == 'complete'
        )
        logging.info("頁面 JavaScript 加載完成")
    except TimeoutException:
        logging.warning("頁面加載超時，繼續嘗試操作")

def check_network(url, timeout=5):
    try:
        response = requests.get(url, timeout=timeout)
        if response.status_code == 200:
            logging.info(f"網絡檢查成功: {url}")
            return True
        else:
            logging.warning(f"網絡檢查失敗，狀態碼: {response.status_code}")
            return False
    except requests.RequestException as e:
        logging.error(f"網絡檢查失敗: {e}")
        return False

def handle_popup(driver, wait):
    try:
        error_div = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'System Error') or contains(@class, 'MuiDialog-container') or contains(@class, 'MuiDialog') and not(@aria-label='menu')]"))
        )
        logging.info("檢測到彈出視窗")
        close_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Close') or contains(text(), 'OK') or contains(text(), 'Cancel') or contains(@class, 'MuiButton') and not(@aria-label='menu')]"))
        )
        wait.until(EC.visibility_of(close_button))
        driver.execute_script("arguments[0].scrollIntoView(true);", close_button)
        time.sleep(0.5)
        close_button.click()
        logging.info("已點擊關閉按鈕")
        WebDriverWait(driver, 3).until(
            EC.invisibility_of_element_located((By.XPATH, "//div[contains(text(), 'System Error') or contains(@class, 'MuiDialog-container') or contains(@class, 'MuiDialog')]"))
        )
        logging.info("彈出視窗已消失")
    except TimeoutException:
        logging.debug("無彈出視窗檢測到")
    except ElementClickInterceptedException as e:
        logging.warning(f"關閉彈出視窗失敗: {str(e)}")
        driver.save_screenshot("popup_close_failure.png")
        with open("popup_close_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)

def cplus_login(driver, wait):
    logging.info("嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/")
    if not check_network("https://cplus.hit.com.hk"):
        logging.error("CPLUS 網站網絡不可用")
        raise Exception("CPLUS 網站網絡不可用")
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    wait_for_page_load(driver)
    logging.info(f"網站已成功打開，當前 URL: {driver.current_url}")

    logging.info("點擊登錄前按鈕...")
    for attempt in range(3):
        try:
            login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button[not(@aria-label='menu')]/span[1]")))
            wait.until(EC.visibility_of(login_button_pre))
            driver.execute_script("arguments[0].scrollIntoView(true);", login_button_pre)
            time.sleep(0.5)
            handle_popup(driver, wait)
            login_button_pre.click()
            logging.info("登錄前按鈕點擊成功")
            break
        except (TimeoutException, ElementClickInterceptedException) as e:
            logging.warning(f"登錄前按鈕點擊失敗 (嘗試 {attempt+1}/3, URL: {driver.current_url}): {str(e)}")
            handle_popup(driver, wait)
            driver.save_screenshot(f"cplus_login_pre_failure_attempt_{attempt+1}.png")
            with open(f"cplus_login_pre_failure_attempt_{attempt+1}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            if attempt < 2:
                time.sleep(0.5)
                continue
            try:
                driver.execute_script("arguments[0].click();", login_button_pre)
                logging.info("登錄前按鈕 JavaScript 點擊成功")
                break
            except Exception as js_e:
                logging.error(f"登錄前按鈕 JavaScript 點擊失敗: {str(js_e)}")
                raise Exception("登錄前按鈕點擊失敗")

    logging.info("輸入 COMPANY CODE...")
    company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    company_code_field.send_keys("CKL")
    logging.info("COMPANY CODE 輸入完成")
    time.sleep(0.5)

    logging.info("輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
    user_id_field.send_keys("KEN")
    logging.info("USER ID 輸入完成")
    time.sleep(0.5)

    logging.info("輸入 PASSWORD...")
    password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
    password_field.send_keys(os.environ.get('SITE_PASSWORD'))
    logging.info("PASSWORD 輸入完成")
    time.sleep(0.5)

    logging.info("點擊 LOGIN 按鈕...")
    for attempt in range(3):
        try:
            login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button[not(@aria-label='menu')]/span[1]")))
            wait.until(EC.visibility_of(login_button))
            driver.execute_script("arguments[0].scrollIntoView(true);", login_button)
            time.sleep(0.5)
            handle_popup(driver, wait)
            login_button.click()
            logging.info("LOGIN 按鈕點擊成功")
            break
        except (TimeoutException, ElementClickInterceptedException) as e:
            logging.warning(f"LOGIN 按鈕點擊失敗 (嘗試 {attempt+1}/3, URL: {driver.current_url}): {str(e)}")
            handle_popup(driver, wait)
            driver.save_screenshot(f"cplus_login_failure_attempt_{attempt+1}.png")
            with open(f"cplus_login_failure_attempt_{attempt+1}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            if attempt < 2:
                time.sleep(0.5)
                continue
            try:
                driver.execute_script("arguments[0].click();", login_button)
                logging.info("LOGIN 按鈕 JavaScript 點擊成功")
                break
            except Exception as js_e:
                logging.error(f"LOGIN 按鈕 JavaScript 點擊失敗: {str(js_e)}")
                raise Exception("LOGIN 按鈕點擊失敗")

def process_cplus_movement(driver, wait, initial_files):
    for page_attempt in range(3):
        logging.info(f"直接前往 Container Movement Log (嘗試 {page_attempt+1}/3)...")
        if not check_network("https://cplus.hit.com.hk"):
            logging.error("CPLUS 網站網絡不可用")
            raise Exception("CPLUS 網站網絡不可用")
        driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
        wait_for_page_load(driver)
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")))
            logging.info("Container Movement Log 頁面加載完成")
        except TimeoutException:
            logging.warning(f"Container Movement Log 頁面加載失敗 (嘗試 {page_attempt+1}/3)，記錄頁面狀態...")
            driver.save_screenshot(f"movement_page_failure_attempt_{page_attempt+1}.png")
            with open(f"movement_page_failure_attempt_{page_attempt+1}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            if page_attempt < 2:
                continue
            raise Exception("Container Movement Log 頁面加載失敗")

        logging.info("點擊 Search...")
        local_initial = initial_files.copy()
        for attempt in range(3):
            try:
                search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button[not(@aria-label='menu')]")))
                wait.until(EC.visibility_of(search_button))
                driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
                time.sleep(0.5)
                handle_popup(driver, wait)
                search_button.click()
                logging.info("Search 按鈕點擊成功")
                break
            except (TimeoutException, ElementClickInterceptedException) as e:
                logging.warning(f"Search 按鈕定位或點擊失敗 (嘗試 {attempt+1}/3, URL: {driver.current_url}): {str(e)}")
                handle_popup(driver, wait)
                driver.save_screenshot(f"movement_search_failure_attempt_{attempt+1}.png")
                with open(f"movement_search_failure_attempt_{attempt+1}.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                if attempt < 2:
                    continue
                try:
                    search_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button[not(@aria-label='menu')]")
                    driver.execute_script("arguments[0].click();", search_button)
                    logging.info("Search 按鈕 JavaScript 點擊成功")
                    break
                except Exception as js_e:
                    logging.error(f"Search 按鈕 JavaScript 點擊失敗: {str(js_e)}")
                    raise Exception("Container Movement Log Search 按鈕點擊失敗")

        logging.info("點擊 Download...")
        for attempt in range(3):
            try:
                download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button[not(@aria-label='menu')]")))
                wait.until(EC.visibility_of(download_button))
                driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
                time.sleep(0.5)
                handle_popup(driver, wait)
                download_button.click()
                logging.info("Download 按鈕點擊成功")
                break
            except (TimeoutException, ElementClickInterceptedException) as e:
                logging.warning(f"Download 按鈕點擊失敗 (嘗試 {attempt+1}/3, URL: {driver.current_url}): {str(e)}")
                handle_popup(driver, wait)
                driver.save_screenshot(f"movement_download_failure_attempt_{attempt+1}.png")
                with open(f"movement_download_failure_attempt_{attempt+1}.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                if attempt < 2:
                    time.sleep(0.5)
                    continue
                try:
                    download_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button[not(@aria-label='menu')]")
                    driver.execute_script("arguments[0].click();", download_button)
                    logging.info("Download 按鈕 JavaScript 點擊成功")
                    break
                except Exception as js_e:
                    logging.error(f"Download 按鈕 JavaScript 點擊失敗: {str(js_e)}")
                    raise Exception("Container Movement Log Download 按鈕點擊失敗")

        new_files = wait_for_new_file(download_dir, local_initial)
        if new_files:
            logging.info(f"Container Movement Log 下載完成，檔案位於: {download_dir}")
            filtered_files = {f for f in new_files if "cntrMoveLog" in f}
            for file in filtered_files:
                logging.info(f"新下載檔案: {file}")
            if not filtered_files:
                logging.warning(f"未下載預期檔案 (cntrMoveLog.xlsx)，記錄頁面狀態...")
                driver.save_screenshot(f"movement_download_failure_attempt_{page_attempt+1}.png")
                raise Exception("Container Movement Log 未下載預期檔案")
            return filtered_files
        else:
            logging.warning(f"Container Movement Log 未觸發新文件下載 (嘗試 {page_attempt+1}/3)，記錄頁面狀態...")
            driver.save_screenshot(f"movement_download_failure_attempt_{page_attempt+1}.png")
            if page_attempt < 2:
                continue
            raise Exception("Container Movement Log 未觸發新文件下載")

def process_cplus_onhand(driver, wait, initial_files):
    logging.info("前往 OnHandContainerList 頁面...")
    if not check_network("https://cplus.hit.com.hk"):
        logging.error("CPLUS 網站網絡不可用")
        raise Exception("CPLUS 網站網絡不可用")
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/OnHandContainerList")
    wait_for_page_load(driver)
    wait.until(EC.presence_of_element_located((By.ID, "root")))
    logging.info("OnHandContainerList 頁面加載完成")

    logging.info("點擊 Search...")
    local_initial = initial_files.copy()
    for attempt in range(5):
        try:
            search_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.MuiPaper-root button.MuiButton-containedPrimary span.MuiButton-label")))
            wait.until(EC.visibility_of(search_button))
            driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
            time.sleep(0.5)
            handle_popup(driver, wait)
            search_button.click()
            logging.info("Search 按鈕點擊成功")
            break
        except (TimeoutException, ElementClickInterceptedException) as e:
            logging.warning(f"Search 按鈕定位或點擊失敗 (嘗試 {attempt+1}/5, URL: {driver.current_url}): {str(e)}")
            handle_popup(driver, wait)
            driver.save_screenshot(f"onhand_search_failure_attempt_{attempt+1}.png")
            with open(f"onhand_search_failure_attempt_{attempt+1}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            if attempt < 4:
                driver.refresh()
                wait_for_page_load(driver)
                time.sleep(0.5)
                continue
            try:
                search_button = driver.find_element(By.XPATH, "//div[contains(@class, 'MuiPaper-root')]//button[contains(text(), 'Search') or contains(@class, 'MuiButton-contained') and not(@aria-label='menu')]")
                driver.execute_script("arguments[0].click();", search_button)
                logging.info("Search 按鈕 JavaScript 點擊成功")
                break
            except Exception as js_e:
                logging.error(f"Search 按鈕 JavaScript 點擊失敗: {str(js_e)}")
                raise Exception("OnHandContainerList Search 按鈕點擊失敗")

    logging.info("點擊 Export...")
    for attempt in range(5):
        try:
            export_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.MuiPaper-root button[aria-label='export'] span.MuiButton-label")))
            wait.until(EC.visibility_of(export_button))
            driver.execute_script("arguments[0].scrollIntoView(true);", export_button)
            time.sleep(0.5)
            handle_popup(driver, wait)
            export_button.click()
            logging.info("Export 按鈕點擊成功")
            break
        except (TimeoutException, ElementClickInterceptedException) as e:
            logging.warning(f"Export 按鈕定位或點擊失敗 (嘗試 {attempt+1}/5, URL: {driver.current_url}): {str(e)}")
            handle_popup(driver, wait)
            driver.save_screenshot(f"onhand_export_failure_attempt_{attempt+1}.png")
            with open(f"onhand_export_failure_attempt_{attempt+1}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            if attempt < 4:
                driver.refresh()
                wait_for_page_load(driver)
                time.sleep(0.5)
                continue
            try:
                export_button = driver.find_element(By.XPATH, "//div[contains(@class, 'MuiPaper-root')]//button[contains(text(), 'Export') or contains(@aria-label, 'export') and not(@aria-label='menu')]")
                driver.execute_script("arguments[0].click();", export_button)
                logging.info("Export 按鈕 JavaScript 點擊成功")
                break
            except Exception as js_e:
                logging.error(f"Export 按鈕 JavaScript 點擊失敗: {str(js_e)}")
                raise Exception("OnHandContainerList Export 按鈕點擊失敗")

    logging.info("點擊 Export as CSV...")
    for attempt in range(3):
        try:
            export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and contains(text(), 'Export as CSV') and not(@aria-label='menu')]")))
            wait.until(EC.visibility_of(export_csv_button))
            driver.execute_script("arguments[0].scrollIntoView(true);", export_csv_button)
            time.sleep(0.5)
            handle_popup(driver, wait)
            export_csv_button.click()
            logging.info("Export as CSV 按鈕點擊成功")
            break
        except (TimeoutException, ElementClickInterceptedException) as e:
            logging.warning(f"Export as CSV 按鈕定位或點擊失敗 (嘗試 {attempt+1}/3, URL: {driver.current_url}): {str(e)}")
            handle_popup(driver, wait)
            driver.save_screenshot(f"onhand_export_csv_failure_attempt_{attempt+1}.png")
            with open(f"onhand_export_csv_failure_attempt_{attempt+1}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            if attempt < 2:
                time.sleep(0.5)
                continue
            try:
                export_csv_button = driver.find_element(By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and contains(text(), 'Export as CSV') and not(@aria-label='menu')]")
                driver.execute_script("arguments[0].click();", export_csv_button)
                logging.info("Export as CSV 按鈕 JavaScript 點擊成功")
                break
            except Exception as js_e:
                logging.error(f"Export as CSV 按鈕 JavaScript 點擊失敗: {str(js_e)}")
                raise Exception("OnHandContainerList Export as CSV 按鈕點擊失敗")

    new_files = wait_for_new_file(download_dir, local_initial)
    if new_files:
        logging.info(f"OnHandContainerList 下載完成，檔案位於: {download_dir}")
        filtered_files = {f for f in new_files if "data_" in f}
        for file in filtered_files:
            logging.info(f"新下載檔案: {file}")
        if not filtered_files:
            logging.warning("未下載預期檔案 (data_*.csv)，記錄頁面狀態...")
            driver.save_screenshot("onhand_download_failure.png")
            with open("onhand_download_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            raise Exception("OnHandContainerList 未下載預期檔案")
        return filtered_files
    else:
        logging.warning("OnHandContainerList 未觸發新文件下載，記錄頁面狀態...")
        driver.save_screenshot("onhand_download_failure.png")
        with open("onhand_download_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("OnHandContainerList 未觸發新文件下載")

def process_cplus_house(driver, wait, initial_files):
    logging.info("前往 Housekeeping Reports 頁面...")
    if not check_network("https://cplus.hit.com.hk"):
        logging.error("CPLUS 網站網絡不可用")
        raise Exception("CPLUS 網站網絡不可用")
    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
    wait_for_page_load(driver)
    wait.until(EC.presence_of_element_located((By.ID, "root")))
    logging.info("Housekeeping Reports 頁面加載完成")

    logging.info("等待表格加載...")
    for attempt in range(5):
        try:
            rows = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table.MuiTable-root tbody tr")))
            if len(rows) == 0 or all(not row.text.strip() for row in rows):
                logging.debug("表格數據空或無效，刷新頁面...")
                driver.refresh()
                wait_for_page_load(driver)
                rows = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table.MuiTable-root tbody tr")))
            if len(rows) < 6:
                logging.warning("表格數據不足，記錄頁面狀態...")
                driver.save_screenshot(f"house_load_failure_attempt_{attempt+1}.png")
                with open(f"house_load_failure_attempt_{attempt+1}.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                if attempt < 4:
                    continue
                raise Exception("Housekeeping Reports 表格數據不足")
            logging.info("表格加載完成")
            break
        except TimeoutException:
            logging.warning(f"表格未加載，嘗試刷新頁面 (嘗試 {attempt+1}/5)...")
            driver.refresh()
            wait_for_page_load(driver)
            if attempt == 4:
                logging.error("表格加載失敗，記錄頁面狀態...")
                driver.save_screenshot("house_load_failure.png")
                with open("house_load_failure.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                raise Exception("Housekeeping Reports 表格加載失敗")

    logging.info("定位並點擊所有 Excel 下載按鈕...")
    local_initial = initial_files.copy()
    new_files = set()
    excel_buttons = driver.find_elements(By.CSS_SELECTOR, "table.MuiTable-root tbody tr td:nth-child(4) button.MuiIconButton-root:not([disabled]):not([aria-label='menu'])")
    button_count = len(excel_buttons)
    logging.info(f"找到 {button_count} 個 Excel 下載按鈕")

    if button_count == 0:
        logging.debug("未找到 Excel 按鈕，嘗試備用定位...")
        excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@aria-label='menu') and contains(@class, 'MuiIconButton-root') and not(@disabled)]")
        button_count = len(excel_buttons)
        logging.info(f"備用定位找到 {button_count} 個 Excel 下載按鈕")

    for idx in range(button_count):
        success = False
        try:
            button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"table.MuiTable-root tbody tr:nth-child({idx+1}) td:nth-child(4) button.MuiIconButton-root:not([disabled]):not([aria-label='menu'])")))
            wait.until(EC.visibility_of(button))
            driver.execute_script("arguments[0].scrollIntoView(true);", button)
            time.sleep(0.5)
            handle_popup(driver, wait)
            button.click()
            logging.info(f"第 {idx+1} 個 Excel 按鈕點擊成功")
            temp_new = wait_for_new_file(download_dir, local_initial)
            if temp_new:
                logging.info(f"第 {idx+1} 個按鈕下載新文件: {', '.join(temp_new)}")
                local_initial.update(temp_new)
                new_files.update(temp_new)
                success = True
            else:
                logging.warning(f"第 {idx+1} 個按鈕未觸發新文件下載")
                driver.save_screenshot(f"house_button_{idx+1}_failure.png")
                with open(f"house_button_{idx+1}_failure.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
        except (TimeoutException, ElementClickInterceptedException) as e:
            logging.error(f"第 {idx+1} 個 Excel 按鈕點擊失敗 (URL: {driver.current_url}): {str(e)}")
            handle_popup(driver, wait)
            driver.save_screenshot(f"house_button_{idx+1}_failure.png")
            with open(f"house_button_{idx+1}_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            try:
                button = driver.find_element(By.CSS_SELECTOR, f"table.MuiTable-root tbody tr:nth-child({idx+1}) td:nth-child(4) button.MuiIconButton-root:not([disabled]):not([aria-label='menu'])")
                driver.execute_script("arguments[0].click();", button)
                logging.info(f"第 {idx+1} 個 Excel 按鈕 JavaScript 點擊成功")
                temp_new = wait_for_new_file(download_dir, local_initial)
                if temp_new:
                    logging.info(f"第 {idx+1} 個按鈕下載新文件: {', '.join(temp_new)}")
                    local_initial.update(temp_new)
                    new_files.update(temp_new)
                    success = True
            except Exception as js_e:
                logging.error(f"第 {idx+1} 個 Excel 按鈕 JavaScript 點擊失敗: {str(js_e)}")
        if not success:
            logging.warning(f"第 {idx+1} 個 Excel 按鈕下載失敗")
    if new_files:
        logging.info(f"Housekeeping Reports 下載完成，共 {len(new_files)} 個文件，預期 {button_count} 個")
        return new_files, len(new_files), button_count
    else:
        logging.warning("Housekeeping Reports 未下載任何文件，記錄頁面狀態...")
        driver.save_screenshot("house_download_failure.png")
        with open("house_download_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("Housekeeping Reports 未下載任何文件")

def process_cplus():
    driver = None
    downloaded_files = set()
    initial_files = set(os.listdir(download_dir))
    house_file_count = 0
    house_button_count = 0
    try:
        driver = webdriver.Chrome(options=get_chrome_options(download_dir))
        logging.info("CPLUS WebDriver 初始化成功")
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        wait = WebDriverWait(driver, WAIT_TIMEOUT)
        cplus_login(driver, wait)
        sections = [
            ('movement', process_cplus_movement),
            ('onhand', process_cplus_onhand),
            ('house', process_cplus_house)
        ]
        for section_name, section_func in sections:
            success = False
            for attempt in range(MAX_RETRIES):
                try:
                    if section_name == 'house':
                        new_files, count, button_count = section_func(driver, wait, initial_files)
                        house_file_count = count
                        house_button_count = button_count
                    else:
                        new_files = section_func(driver, wait, initial_files)
                    downloaded_files.update(new_files)
                    initial_files.update(new_files)
                    success = True
                    break
                except Exception as e:
                    logging.error(f"{section_name} 嘗試 {attempt+1}/{MAX_RETRIES} 失敗: {str(e)}")
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(1)
                    else:
                        break  # 連續失敗後跳出，減少等待時間
            if not success:
                logging.error(f"{section_name} 經過 {MAX_RETRIES} 次嘗試失敗")
        return downloaded_files, house_file_count, house_button_count, driver
    except Exception as e:
        logging.error(f"CPLUS 總錯誤: {str(e)}")
        return downloaded_files, house_file_count, house_button_count, driver
    finally:
        try:
            if driver:
                logging.info("嘗試登出...")
                logout_menu_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button[not(@aria-label='menu')]/span[1]")))
                wait.until(EC.visibility_of(logout_menu_button))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_menu_button)
                time.sleep(0.5)
                handle_popup(driver, wait)
                logout_menu_button.click()
                logging.info("用戶菜單點擊成功")
                logout_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Logout') or contains(text(), 'Sign out') and not(@aria-label='menu')]")))
                wait.until(EC.visibility_of(logout_option))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_option)
                time.sleep(0.5)
                handle_popup(driver, wait)
                logout_option.click()
                logging.info("Logout 選項點擊成功")
                try:
                    close_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='logout']/div[3]/div/div[3]/button[not(@aria-label='menu')]/span[1]")))
                    wait.until(EC.visibility_of(close_button))
                    driver.execute_script("arguments[0].scrollIntoView(true);", close_button)
                    time.sleep(0.5)
                    handle_popup(driver, wait)
                    close_button.click()
                    logging.info("Logout 後 CLOSE 按鈕點擊成功")
                except TimeoutException:
                    logging.warning("Logout 後無 CLOSE 按鈕，跳過")
        except (TimeoutException, ElementClickInterceptedException) as e:
            logging.error(f"登出失敗 (URL: {driver.current_url}): {str(e)}")
            handle_popup(driver, wait)
            driver.save_screenshot("cplus_logout_failure.png")
            with open("cplus_logout_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)

def barge_login(driver, wait):
    logging.info("嘗試打開網站 https://barge.oneport.com/login...")
    if not check_network("https://barge.oneport.com"):
        logging.error("Barge 網站網絡不可用")
        raise Exception("Barge 網站網絡不可用")
    driver.get("https://barge.oneport.com/login")
    logging.info(f"網站已成功打開，當前 URL: {driver.current_url}")

    logging.info("輸入 COMPANY ID...")
    company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Company ID' or contains(@id, 'mat-input-0')]")))
    company_id_field.send_keys("CKL")
    logging.info("COMPANY ID 輸入完成")
    time.sleep(0.5)

    logging.info("輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='User ID' or contains(@id, 'mat-input-1')]")
    user_id_field.send_keys("barge")
    logging.info("USER ID 輸入完成")
    time.sleep(0.5)

    logging.info("輸入 PW...")
    password_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Password' or contains(@id, 'mat-input-2')]")
    password_field.send_keys(os.environ.get('BARGE_PASSWORD', '123456'))
    logging.info("PW 輸入完成")
    time.sleep(0.5)

    logging.info("點擊 LOGIN 按鈕...")
    login_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'LOGIN') or contains(@class, 'mat-raised-button') and not(@aria-label='menu')]")))
    wait.until(EC.visibility_of(login_button_barge))
    driver.execute_script("arguments[0].scrollIntoView(true);", login_button_barge)
    time.sleep(0.5)
    handle_popup(driver, wait)
    login_button_barge.click()
    logging.info("LOGIN 按鈕點擊成功")

def process_barge_download(driver, wait, initial_files):
    logging.info("直接前往 https://barge.oneport.com/downloadReport...")
    if not check_network("https://barge.oneport.com"):
        logging.error("Barge 網站網絡不可用")
        raise Exception("Barge 網站網絡不可用")
    driver.get("https://barge.oneport.com/downloadReport")
    wait_for_page_load(driver)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    logging.info("downloadReport 頁面加載完成")

    logging.info("選擇 Report Type...")
    for attempt in range(3):
        try:
            report_type_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-form-field[.//mat-label[contains(text(), 'Report Type')]]//div[contains(@class, 'mat-select-trigger') and not(@aria-label='menu')]")))
            wait.until(EC.visibility_of(report_type_trigger))
            driver.execute_script("arguments[0].scrollIntoView(true);", report_type_trigger)
            time.sleep(0.5)
            handle_popup(driver, wait)
            report_type_trigger.click()
            logging.info("Report Type 選擇開始")
            break
        except (TimeoutException, ElementClickInterceptedException) as e:
            logging.warning(f"Report Type 按鈕點擊失敗 (嘗試 {attempt+1}/3, URL: {driver.current_url}): {str(e)}")
            handle_popup(driver, wait)
            driver.save_screenshot(f"barge_report_type_failure_attempt_{attempt+1}.png")
            with open(f"barge_report_type_failure_attempt_{attempt+1}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            if attempt < 2:
                time.sleep(0.5)
                continue
            try:
                report_type_trigger = driver.find_element(By.XPATH, "//mat-form-field[.//mat-label[contains(text(), 'Report Type')]]//div[contains(@class, 'mat-select-trigger') and not(@aria-label='menu')]")
                driver.execute_script("arguments[0].click();", report_type_trigger)
                logging.info("Report Type 按鈕 JavaScript 點擊成功")
                break
            except Exception as js_e:
                logging.error(f"Report Type 按鈕 JavaScript 點擊失敗: {str(js_e)}")
                raise Exception("Report Type 按鈕點擊失敗")

    logging.info("點擊 Container Detail...")
    for attempt in range(3):
        try:
            container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-option//span[contains(text(), 'Container Detail') and not(@aria-label='menu')]")))
            wait.until(EC.visibility_of(container_detail_option))
            driver.execute_script("arguments[0].scrollIntoView(true);", container_detail_option)
            time.sleep(0.5)
            handle_popup(driver, wait)
            container_detail_option.click()
            logging.info("Container Detail 點擊成功")
            break
        except (TimeoutException, ElementClickInterceptedException) as e:
            logging.warning(f"Container Detail 按鈕點擊失敗 (嘗試 {attempt+1}/3, URL: {driver.current_url}): {str(e)}")
            handle_popup(driver, wait)
            driver.save_screenshot(f"barge_container_detail_failure_attempt_{attempt+1}.png")
            with open(f"barge_container_detail_failure_attempt_{attempt+1}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            if attempt < 2:
                time.sleep(0.5)
                continue
            try:
                container_detail_option = driver.find_element(By.XPATH, "//mat-option//span[contains(text(), 'Container Detail') and not(@aria-label='menu')]")
                driver.execute_script("arguments[0].click();", container_detail_option)
                logging.info("Container Detail 按鈕 JavaScript 點擊成功")
                break
            except Exception as js_e:
                logging.error(f"Container Detail 按鈕 JavaScript 點擊失敗: {str(js_e)}")
                raise Exception("Container Detail 按鈕點擊失敗")

    logging.info("點擊 Download...")
    local_initial = initial_files.copy()
    for attempt in range(3):
        try:
            download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Download'] and not(@aria-label='menu')]")))
            wait.until(EC.visibility_of(download_button_barge))
            driver.execute_script("arguments[0].scrollIntoView(true);", download_button_barge)
            time.sleep(0.5)
            handle_popup(driver, wait)
            download_button_barge.click()
            logging.info("Download 按鈕點擊成功")
            break
        except (TimeoutException, ElementClickInterceptedException) as e:
            logging.warning(f"Download 按鈕點擊失敗 (嘗試 {attempt+1}/3, URL: {driver.current_url}): {str(e)}")
            handle_popup(driver, wait)
            driver.save_screenshot(f"barge_download_failure_attempt_{attempt+1}.png")
            with open(f"barge_download_failure_attempt_{attempt+1}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            if attempt < 2:
                time.sleep(0.5)
                continue
            try:
                download_button_barge = driver.find_element(By.XPATH, "//button[span[text()='Download'] and not(@aria-label='menu')]")
                driver.execute_script("arguments[0].click();", download_button_barge)
                logging.info("Download 按鈕 JavaScript 點擊成功")
                break
            except Exception as js_e:
                logging.error(f"Download 按鈕 JavaScript 點擊失敗: {str(js_e)}")
                raise Exception("Container Detail Download 按鈕點擊失敗")

    new_files = wait_for_new_file(download_dir, local_initial)
    if new_files:
        logging.info(f"Container Detail 下載完成，檔案位於: {download_dir}")
        filtered_files = {f for f in new_files if "ContainerDetailReport" in f}
        for file in filtered_files:
            logging.info(f"新下載檔案: {file}")
        if not filtered_files:
            logging.warning("未下載預期檔案 (ContainerDetailReport*.csv)，記錄頁面狀態...")
            driver.save_screenshot("barge_download_failure.png")
            raise Exception("Container Detail 未下載預期檔案")
        return filtered_files
    else:
        logging.warning("Container Detail 未觸發新文件下載，記錄頁面狀態...")
        driver.save_screenshot("barge_download_failure.png")
        raise Exception("Container Detail 未觸發新文件下載")

def process_barge():
    driver = None
    downloaded_files = set()
    initial_files = set(os.listdir(download_dir))
    try:
        driver = webdriver.Chrome(options=get_chrome_options(download_dir))
        logging.info("Barge WebDriver 初始化成功")
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        wait = WebDriverWait(driver, WAIT_TIMEOUT)
        barge_login(driver, wait)
        success = False
        for attempt in range(MAX_RETRIES):
            try:
                new_files = process_barge_download(driver, wait, initial_files)
                downloaded_files.update(new_files)
                initial_files.update(new_files)
                success = True
                break
            except Exception as e:
                logging.error(f"Barge 下載嘗試 {attempt+1}/{MAX_RETRIES} 失敗: {str(e)}")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(1)
                else:
                    break  # 連續失敗後跳出
        if not success:
            logging.error(f"Barge 下載經過 {MAX_RETRIES} 次嘗試失敗")
        return downloaded_files, driver
    except Exception as e:
        logging.error(f"Barge 總錯誤: {str(e)}")
        return downloaded_files, driver
    finally:
        try:
            if driver:
                logging.info("點擊工具欄進行登出...")
                logout_toolbar_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1][not(@aria-label='menu')]")))
                wait.until(EC.visibility_of(logout_toolbar_barge))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar_barge)
                time.sleep(0.5)
                handle_popup(driver, wait)
                logout_toolbar_barge.click()
                logging.info("工具欄點擊成功")
                logout_span_xpath = "//div[contains(@class, 'mat-menu-panel')]//button//span[contains(text(), 'Logout') and not(@aria-label='menu')]"
                logout_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, logout_span_xpath)))
                wait.until(EC.visibility_of(logout_button_barge))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                time.sleep(0.5)
                handle_popup(driver, wait)
                logout_button_barge.click()
                logging.info("Logout 選項點擊成功")
        except Exception as e:
            logging.error(f"登出失敗 (URL: {driver.current_url}): {str(e)}")
            driver.save_screenshot("barge_logout_failure.png")
            with open("barge_logout_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)

def main():
    load_dotenv()
    clear_download_dirs()
    cplus_files = set()
    house_file_count = [0]
    house_button_count = [0]
    barge_files = set()
    cplus_driver = None
    barge_driver = None

    cplus_result = process_cplus()
    cplus_files.update(cplus_result[0])
    house_file_count[0] = cplus_result[1]
    house_button_count[0] = cplus_result[2]
    cplus_driver = cplus_result[3]
    logging.info("CPLUS 處理完成")

    barge_result = process_barge()
    barge_files.update(barge_result[0])
    barge_driver = barge_result[1]
    logging.info("Barge 處理完成")

    logging.info("檢查所有下載文件...")
    downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
    logging.info(f"總下載文件: {len(downloaded_files)} 個")
    for file in downloaded_files:
        logging.info(f"找到檔案: {file}")

    required_patterns = {'movement': 'cntrMoveLog', 'onhand': 'data_', 'barge': 'ContainerDetailReport'}
    housekeep_prefixes = ['IE2_', 'DM1C_', 'IA17_', 'GA1_', 'IA5_', 'IA15_']
    has_required = all(any(pattern in f for f in downloaded_files) for pattern in required_patterns.values())
    house_files = [f for f in downloaded_files if any(p in f for p in housekeep_prefixes)]
    house_download_count = len(house_files)
    house_ok = (house_button_count[0] == 0) or (house_download_count >= house_button_count[0])

    if has_required and house_ok:
        logging.info("所有必須文件齊全，開始發送郵件...")
        try:
            smtp_server = os.environ.get('SMTP_SERVER', 'smtp.zoho.com')
            smtp_port = int(os.environ.get('SMTP_PORT', 587))
            sender_email = os.environ['ZOHO_EMAIL']
            sender_password = os.environ['ZOHO_PASSWORD']
            receiver_emails = os.environ.get('RECEIVER_EMAILS', 'paklun@ckline.com.hk').split(',')
            cc_emails = os.environ.get('CC_EMAILS', '').split(',') if os.environ.get('CC_EMAILS') else []
            dry_run = os.environ.get('DRY_RUN', 'False').lower() == 'true'
            if dry_run:
                logging.info("Dry run 模式：只打印郵件內容，不發送。")
            house_report_names = ["REEFER CONTAINER MONITOR REPORT", "CONTAINER DAMAGE REPORT (LINE) ENTRY GATE + EXIT GATE", "CONTAINER LIST (ON HAND)", "CY - GATELOG", "CONTAINER LIST (DAMAGED)", "ACTIVE REEFER CONTAINER ON HAND LIST"]
            house_status = ['✓' if [f for f in house_files if p in f] else '-' for p in housekeep_prefixes]
            house_file_names = [', '.join([f for f in house_files if p in f]) if [f for f in house_files if p in f] else 'N/A' for p in housekeep_prefixes]
            body_html = f"""
            <html><body><p>Attached are the daily reports downloaded from CPLUS and Barge. Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <table border="1" style="border-collapse: collapse; width: 100%;"><thead><tr><th>Category</th><th>Report</th><th>File Names</th><th>Status</th></tr></thead><tbody>
            <tr><td rowspan="8">CPLUS</td><td>Container Movement</td><td>{', '.join([f for f in downloaded_files if 'cntrMoveLog' in f]) or 'N/A'}</td><td>{'✓' if any('cntrMoveLog' in f for f in downloaded_files) else '-'}</td></tr>
            <tr><td>OnHandContainerList</td><td>{', '.join([f for f in downloaded_files if 'data_' in f]) or 'N/A'}</td><td>{'✓' if any('data_' in f for f in downloaded_files) else '-'}</td></tr>
            <tr><td>{house_report_names[0]}</td><td>{house_file_names[0]}</td><td>{house_status[0]}</td></tr>
            <tr><td>{house_report_names[1]}</td><td>{house_file_names[1]}</td><td>{house_status[1]}</td></tr>
            <tr><td>{house_report_names[2]}</td><td>{house_file_names[2]}</td><td>{house_status[2]}</td></tr>
            <tr><td>{house_report_names[3]}</td><td>{house_file_names[3]}</td><td>{house_status[3]}</td></tr>
            <tr><td>{house_report_names[4]}</td><td>{house_file_names[4]}</td><td>{house_status[4]}</td></tr>
            <tr><td>{house_report_names[5]}</td><td>{house_file_names[5]}</td><td>{house_status[5]}</td></tr>
            <tr><td rowspan="1">BARGE</td><td>Container Detail</td><td>{', '.join([f for f in downloaded_files if 'ContainerDetailReport' in f]) or 'N/A'}</td><td>{'✓' if any('ContainerDetailReport' in f for f in downloaded_files) else '-'}</td></tr>
            <tr><td colspan="2"><strong>TOTAL</strong></td><td><strong>{len(downloaded_files)} files attached</strong></td><td><strong>{len(downloaded_files)}</strong></td></tr>
            </tbody></table></body></html>
            """
            msg = MIMEMultipart('alternative')
            msg['From'] = sender_email
            msg['To'] = ', '.join(receiver_emails)
            if cc_emails:
                msg['Cc'] = ', '.join(cc_emails)
            msg['Subject'] = f"[TESTING] HIT DAILY {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            msg.attach(MIMEText(body_html, 'html'))
            plain_text = body_html.replace('<br>', '\n').replace('<table>', '').replace('</table>', '').replace('<tr>', '\n').replace('<td>', ' | ').replace('</td>', '').replace('<th>', ' | ').replace('</th>', '').strip()
            msg.attach(MIMEText(plain_text, 'plain'))
            for file in downloaded_files:
                file_path = os.path.join(download_dir, file)
                if os.path.exists(file_path):
                    attachment = MIMEBase('application', 'octet-stream')
                    attachment.set_payload(open(file_path, 'rb').read())
                    encoders.encode_base64(attachment)
                    attachment.add_header('Content-Disposition', f'attachment; filename={file}')
                    msg.attach(attachment)
                else:
                    logging.warning(f"附件不存在: {file_path}")
            if not os.environ.get('DRY_RUN', 'False').lower() == 'true':
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(sender_email, sender_password)
                all_receivers = receiver_emails + cc_emails
                server.sendmail(sender_email, all_receivers, msg.as_string())
                server.quit()
                logging.info("郵件發送成功!")
            else:
                logging.info(f"模擬發送郵件：\nFrom: {sender_email}\nTo: {msg['To']}\nCc: {msg.get('Cc', '')}\nSubject: {msg['Subject']}\nBody: {body_html}")
        except KeyError as ke:
            logging.error(f"缺少環境變量: {ke}")
        except smtplib.SMTPAuthenticationError:
            logging.error("SMTP 認證失敗：檢查用戶名/密碼")
        except smtplib.SMTPConnectError:
            logging.error("SMTP 連接失敗：檢查伺服器/端口")
        except Exception as e:
            logging.error(f"郵件發送失敗: {str(e)}")
    else:
        logging.warning(f"文件不齊全: 缺少必須文件 (has_required={has_required}) 或 House文件不足 (download={house_download_count}, button={house_button_count[0]})")

    if cplus_driver:
        cplus_driver.quit()
        logging.info("CPLUS WebDriver 關閉")
    if barge_driver:
        barge_driver.quit()
        logging.info("Barge WebDriver 關閉")
    logging.info("腳本完成")

if __name__ == "__main__":
    setup_environment()
    main()
