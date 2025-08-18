import os
import time
import shutil
import subprocess
import threading
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
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
import logging
from dotenv import load_dotenv

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

cplus_download_dir = os.path.abspath("downloads_cplus")
barge_download_dir = os.path.abspath("downloads_barge")
MAX_RETRIES = 1
DOWNLOAD_TIMEOUT = 20

def clear_download_dirs():
    for dir_path in [cplus_download_dir, barge_download_dir]:
        if os.path.exists(dir_path):
            shutil.rmtree(dir_path)
        os.makedirs(dir_path)
        logging.info(f"創建下載目錄: {dir_path}")

def setup_environment():
    os.environ.pop('TZ', None)
    try:
        result = subprocess.run(['which', 'chromium-browser'], capture_output=True, text=True)
        if result.returncode != 0:
            subprocess.run(['sudo', 'apt-get', 'update', '-qq'], check=True)
            subprocess.run(['sudo', 'apt-get', 'install', '-y', 'chromium-browser', 'chromium-chromedriver'], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            logging.info("Chromium 及 ChromeDriver 已安裝")
        else:
            logging.info("Chromium 及 ChromeDriver 已存在，跳過安裝")

        result = subprocess.run(['pip', 'show', 'selenium'], capture_output=True, text=True)
        if "selenium" not in result.stdout or "webdriver-manager" not in subprocess.run(['pip', 'show', 'webdriver-manager'], capture_output=True, text=True).stdout:
            subprocess.run(['pip', 'install', 'selenium', 'webdriver-manager'], check=True)
            logging.info("Selenium 及 WebDriver Manager 已安裝")
        else:
            logging.info("Selenium 及 WebDriver Manager 已存在，跳過安裝")
    except subprocess.CalledProcessError as e:
        logging.error(f"環境準備失敗: {e}")
        raise

def get_chrome_options(download_dir):
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--disable-popup-blocking')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--no-first-run')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_argument('--enable-logging')
    chrome_options.add_argument('--log-level=0')
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": False
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.binary_location = '/usr/bin/chromium-browser'
    return chrome_options

def wait_for_new_file(download_dir, initial_files, timeout=DOWNLOAD_TIMEOUT):
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_files = set(f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx')))
        new_files = current_files - initial_files
        new_files = {f for f in new_files if os.path.getsize(os.path.join(download_dir, f)) > 10240}
        if new_files:
            return new_files
        time.sleep(0.5)
    return set()

def handle_popup(driver, wait):
    try:
        error_div = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'System Error') or contains(text(), 'Loading') or contains(@class, 'error') or contains(@class, 'popup')]")))
        logging.info("檢測到 popup（System Error 或 Loading）")
        close_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Close') or contains(text(), 'OK') or contains(@class, 'close')]")))
        driver.execute_script("arguments[0].click();", close_button)
        logging.info("已關閉 popup")
        WebDriverWait(driver, 5).until(EC.invisibility_of_element_located((By.XPATH, "//div[contains(text(), 'System Error') or contains(text(), 'Loading') or contains(@class, 'error') or contains(@class, 'popup')]")))
    except TimeoutException:
        logging.debug("無 popup 檢測到")

    # 檢查二因子認証或錯誤提示
    try:
        two_factor_popup = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Two Factor Authentication')]")))
        logging.error("CPLUS: 檢測到二因子認証彈窗，登錄需要手動配置")
        driver.save_screenshot("two_factor_popup.png")
        with open("two_factor_popup.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: 檢測到二因子認証，需手動處理")
    except TimeoutException:
        pass
    try:
        error_message = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'System Error')]")))
        logging.error(f"CPLUS: 檢測到系統錯誤: {error_message.text}")
        driver.save_screenshot("system_error.png")
        with open("system_error.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: 檢測到系統錯誤")
    except TimeoutException:
        pass

def cplus_login(driver, wait):
    logging.info("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/")
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    logging.info(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}")
    time.sleep(0.2)

    logging.info("CPLUS: 點擊登錄前按鈕...")
    login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
    ActionChains(driver).move_to_element(login_button_pre).click().perform()
    logging.info("CPLUS: 登錄前按鈕點擊成功")
    time.sleep(0.2)

    logging.info("CPLUS: 輸入 COMPANY CODE...")
    company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    company_code_field.send_keys("CKL")
    logging.info("CPLUS: COMPANY CODE 輸入完成")
    time.sleep(0.2)

    logging.info("CPLUS: 輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
    user_id_field.send_keys("KEN")
    logging.info("CPLUS: USER ID 輸入完成")
    time.sleep(0.2)

    logging.info("CPLUS: 輸入 PASSWORD...")
    password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
    password_field.send_keys(os.environ.get('SITE_PASSWORD'))
    logging.info("CPLUS: PASSWORD 輸入完成")
    time.sleep(0.2)

    logging.info("CPLUS: 點擊 LOGIN 按鈕...")
    login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
    ActionChains(driver).move_to_element(login_button).click().perform()
    logging.info("CPLUS: LOGIN 按鈕點擊成功")
    try:
        # 先檢查錯誤或 2FA
        handle_popup(driver, wait)
        # 然後檢查登錄成功
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'KEN (CKL)')]")), 5)
        logging.info("CPLUS: 頁面加載完成，登錄成功")
    except TimeoutException:
        logging.error("CPLUS: 登錄失敗，未找到 'KEN (CKL)'")
        driver.save_screenshot("login_failure.png")
        with open("login_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        try:
            browser_logs = driver.get_log('browser')
            with open("browser_log.txt", "w", encoding="utf-8") as log_file:
                for entry in browser_logs:
                    log_file.write(f"{entry['level']}: {entry['message']}\n")
            logging.info("CPLUS: 瀏覽器日志已保存至 browser_log.txt")
        except Exception as log_error:
            logging.warning(f"CPLUS: 無法獲取瀏覽器日志: {str(log_error)}")
        raise Exception("CPLUS: 登錄失敗")

def process_cplus_movement(driver, wait, initial_files):
    logging.info("CPLUS: 直接前往 Container Movement Log...")
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")))
    logging.info("CPLUS: Container Movement Log 頁面加載完成")

    logging.info("CPLUS: 等待表單字段加載...")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")))
    logging.info("CPLUS: 表單字段加載完成")

    logging.info("CPLUS: 點擊 Search...")
    local_initial = initial_files.copy()
    try:
        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")), 10)
        ActionChains(driver).move_to_element(search_button).click().perform()
        logging.info("CPLUS: Search 按鈕點擊成功")
    except TimeoutException:
        logging.warning("CPLUS: Search 按鈕超時，刷新頁面...")
        driver.refresh()
        time.sleep(0.5)
        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")), 10)
        ActionChains(driver).move_to_element(search_button).click().perform()
        logging.info("CPLUS: Search 按鈕點擊成功 (刷新後)")

    logging.info("CPLUS: 點擊 Download...")
    for attempt in range(2):
        try:
            download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button")), 10)
            ActionChains(driver).move_to_element(download_button).click().perform()
            logging.info("CPLUS: Download 按鈕點擊成功")
            time.sleep(0.5)
            try:
                driver.execute_script("arguments[0].click();", download_button)
                logging.debug("CPLUS: Download 按鈕 JavaScript 點擊成功")
            except Exception as js_e:
                logging.debug(f"CPLUS: Download 按鈕 JavaScript 點擊失敗: {str(js_e)}")
            time.sleep(0.5)
            break
        except Exception as e:
            logging.debug(f"CPLUS: Download 按鈕點擊失敗 (嘗試 {attempt+1}/2): {str(e)}")
            driver.save_screenshot("movement_download_failure.png")
            with open("movement_download_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            time.sleep(0.5)
    else:
        raise Exception("CPLUS: Container Movement Log Download 按鈕點擊失敗")

    new_files = wait_for_new_file(cplus_download_dir, local_initial)
    if new_files:
        logging.info(f"CPLUS: Container Movement Log 下載完成，檔案位於: {cplus_download_dir}")
        filtered_files = {f for f in new_files if "cntrMoveLog" in f}
        for file in filtered_files:
            logging.info(f"CPLUS: 新下載檔案: {file}")
        if not filtered_files:
            logging.warning("CPLUS: 未下載預期檔案 (cntrMoveLog.xlsx)，記錄頁面狀態...")
            driver.save_screenshot("movement_download_failure.png")
            with open("movement_download_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            raise Exception("CPLUS: Container Movement Log 未下載預期檔案")
        return filtered_files
    else:
        logging.warning("CPLUS: Container Movement Log 未觸發新文件下載，記錄頁面狀態...")
        driver.save_screenshot("movement_download_failure.png")
        with open("movement_download_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: Container Movement Log 未觸發新文件下載")

def process_cplus_onhand(driver, wait, initial_files):
    logging.info("CPLUS: 前往 OnHandContainerList 頁面...")
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/OnHandContainerList")
    time.sleep(0.5)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    wait.until(EC.presence_of_element_located((By.XPATH, "//form//input")))
    logging.info("CPLUS: OnHandContainerList 頁面加載完成")

    logging.info("CPLUS: 等待表單字段加載...")
    wait.until(EC.presence_of_element_located((By.XPATH, "//form//input")))
    logging.info("CPLUS: 表單字段加載完成")

    logging.info("CPLUS: 點擊 Search...")
    local_initial = initial_files.copy()
    try:
        search_button_onhand = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")), 10)
        try:
            ActionChains(driver).move_to_element(search_button_onhand).click().perform()
            logging.info("CPLUS: Search 按鈕點擊成功")
        except ElementClickInterceptedException:
            logging.debug("CPLUS: Search 按鈕點擊被攔截，使用 JavaScript 點擊")
            driver.execute_script("arguments[0].click();", search_button_onhand)
            logging.info("CPLUS: Search 按鈕 JavaScript 點擊成功")
    except TimeoutException:
        logging.warning("CPLUS: Search 按鈕超時，刷新頁面...")
        driver.refresh()
        time.sleep(0.5)
        search_button_onhand = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")), 10)
        try:
            ActionChains(driver).move_to_element(search_button_onhand).click().perform()
            logging.info("CPLUS: Search 按鈕點擊成功 (刷新後)")
        except ElementClickInterceptedException:
            logging.debug("CPLUS: Search 按鈕點擊被攔截，使用 JavaScript 點擊")
            driver.execute_script("arguments[0].click();", search_button_onhand)
            logging.info("CPLUS: Search 按鈕 JavaScript 點擊成功 (刷新後)")

    time.sleep(0.2)

    logging.info("CPLUS: 點擊 Export...")
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button")), 10)
    ActionChains(driver).move_to_element(export_button).click().perform()
    logging.info("CPLUS: Export 按鈕點擊成功")
    time.sleep(0.2)

    logging.info("CPLUS: 點擊 Export as CSV...")
    export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']")), 10)
    ActionChains(driver).move_to_element(export_csv_button).click().perform()
    logging.info("CPLUS: Export as CSV 按鈕點擊成功")
    time.sleep(0.2)

    new_files = wait_for_new_file(cplus_download_dir, local_initial)
    if new_files:
        logging.info(f"CPLUS: OnHandContainerList 下載完成，檔案位於: {cplus_download_dir}")
        filtered_files = {f for f in new_files if "data_" in f}
        for file in filtered_files:
            logging.info(f"CPLUS: 新下載檔案: {file}")
        if not filtered_files:
            logging.warning("CPLUS: 未下載預期檔案 (data_*.csv)，記錄頁面狀態...")
            driver.save_screenshot("onhand_download_failure.png")
            with open("onhand_download_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            raise Exception("CPLUS: OnHandContainerList 未下載預期檔案")
        return filtered_files
    else:
        logging.warning("CPLUS: OnHandContainerList 未觸發新文件下載，記錄頁面狀態...")
        driver.save_screenshot("onhand_download_failure.png")
        with open("onhand_download_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: OnHandContainerList 未觸發新文件下載")

def process_cplus_house(driver, wait, initial_files):
    logging.info("CPLUS: 前往 Housekeeping Reports 頁面...")
    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
    time.sleep(0.5)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    logging.info("CPLUS: Housekeeping Reports 頁面加載完成")

    logging.info("CPLUS: 等待表格加載...")
    start_time = time.time()
    try:
        WebDriverWait(driver, 15).until_not(EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'spinner') or contains(@class, 'loading')]")))
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]")))
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr[td[contains(text(), 'CY - GATELOG')]]")))
        rows = WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
        if len(rows) == 0:
            logging.warning("CPLUS: 無記錄，嘗試刷新...")
            driver.refresh()
            time.sleep(0.5)
            WebDriverWait(driver, 15).until_not(EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'spinner') or contains(@class, 'loading')]")))
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]")))
            rows = WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
        logging.info(f"CPLUS: 找到 {len(rows)} 個報告行，耗時 {time.time() - start_time:.1f} 秒")
    except TimeoutException:
        logging.warning("CPLUS: 表格加載失敗，嘗試備用 CSS 定位...")
        WebDriverWait(driver, 15).until_not(EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'spinner') or contains(@class, 'loading')]")))
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".MuiTable-root")))
        rows = WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".MuiTable-root tbody tr")))
        if len(rows) == 0:
            logging.error("CPLUS: 表格加載完全失敗，跳過 Housekeeping Reports")
            driver.save_screenshot("house_load_failure.png")
            with open("house_load_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            return set(), 0, 0
        logging.info(f"CPLUS: 找到 {len(rows)} 個報告行 (備用定位)，耗時 {time.time() - start_time:.1f} 秒")

    if time.time() - start_time > 30:
        logging.warning("CPLUS: Housekeeping Reports 加載時間過長，跳過")
        driver.save_screenshot("house_load_timeout.png")
        return set(), 0, 0

    time.sleep(0.2)
    logging.info("CPLUS: 定位並點擊所有 Excel 下載按鈕...")
    local_initial = initial_files.copy()
    new_files = set()
    all_downloaded_files = set()
    excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[contains(., 'Excel') or contains(@class, 'download')]")
    if not excel_buttons:
        excel_buttons = driver.find_elements(By.CSS_SELECTOR, ".MuiTable-root tbody tr td:nth-child(4) button")
    button_count = len(excel_buttons)
    if button_count == 0:
        logging.error("CPLUS: 未找到任何 Excel 下載按鈕，記錄頁面狀態...")
        driver.save_screenshot("house_button_failure.png")
        with open("house_button_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: Housekeeping Reports 未找到下載按鈕")
    logging.info(f"CPLUS: 找到 {button_count} 個 Excel 下載按鈕")

    report_file_mapping = []
    failed_buttons = []
    for idx in range(button_count):
        success = False
        for retry_count in range(1, 2):
            try:
                button_xpath = f"(//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button)[{idx+1}]"
                button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", button)
                time.sleep(0.2)

                try:
                    report_name = driver.find_element(By.XPATH, f"(//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[3])[{idx+1}]").text
                    logging.info(f"CPLUS: 準備點擊第 {idx+1} 個 Excel 按鈕，報告名稱: {report_name}")
                except:
                    report_name = f"未知報告 {idx+1}"

                driver.execute_script("arguments[0].click();", button)
                logging.info(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕 JS 點擊成功 (重試 {retry_count})")
                time.sleep(0.2)
                handle_popup(driver, wait)

                try:
                    WebDriverWait(driver, 5).until(EC.invisibility_of_element_located((By.XPATH, "//*[contains(@class, 'spinner') or contains(@class, 'loading')]")))
                    logging.debug("CPLUS: 加載 spinner 已消失")
                except TimeoutException:
                    logging.debug("CPLUS: 無 spinner 或超時，繼續")

                ActionChains(driver).move_to_element(button).pause(0.2).click().perform()
                logging.info(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕 ActionChains 點擊成功 (重試 {retry_count})")
                time.sleep(0.2)
                handle_popup(driver, wait)

                temp_new = wait_for_new_file(cplus_download_dir, local_initial)
                if temp_new:
                    temp_new = {f for f in temp_new if not any(s in f for s in [' (1)', ' (2)']) and f not in all_downloaded_files}
                    if temp_new:
                        all_downloaded_files.update(temp_new)
                        report_file_mapping.append((report_name, ', '.join(temp_new)))
                        local_initial.update(temp_new)
                        new_files.update(temp_new)
                        success = True
                        break
                    else:
                        logging.warning(f"CPLUS: 第 {idx+1} 個檢測到重複檔，忽略")
                else:
                    logging.warning(f"CPLUS: 第 {idx+1} 個未觸發新文件 (重試 {retry_count})")
                    driver.save_screenshot(f"house_button_{idx+1}_failure_{retry_count}.png")
                    with open(f"house_button_{idx+1}_failure_{retry_count}.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
            except (TimeoutException, NoSuchElementException, ElementClickInterceptedException, WebDriverException) as e:
                logging.error(f"CPLUS: 第 {idx+1} 個失敗: {str(e)}")
                driver.save_screenshot(f"house_button_{idx+1}_failure_{retry_count}.png")
                if "Read timed out" in str(e):
                    logging.warning("CPLUS: 網絡超時，嘗試重啟 WebDriver...")
                    driver.quit()
                    driver = webdriver.Chrome(options=get_chrome_options(cplus_download_dir))
                    wait = WebDriverWait(driver, 10)
                    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
                    time.sleep(0.5)
                    continue
        if not success:
            failed_buttons.append(idx)
            report_file_mapping.append((report_name, "N/A"))

    if new_files:
        logging.info(f"CPLUS: Housekeeping Reports 下載完成，共 {len(new_files)} 個文件，找到 {button_count} 個按鈕")
        for report, files in report_file_mapping:
            logging.info(f"報告: {report}, 文件: {files}")
        if failed_buttons:
            logging.warning(f"CPLUS: {len(failed_buttons)} 個失敗: {failed_buttons}")
    else:
        logging.warning("CPLUS: Housekeeping Reports 下載失敗，但繼續執行")

    return new_files, len(new_files), button_count

def process_cplus():
    driver = None
    downloaded_files = set()
    initial_files = set(os.listdir(cplus_download_dir))
    house_file_count = 0
    house_button_count = 0
    try:
        driver = webdriver.Chrome(options=get_chrome_options(cplus_download_dir))
        logging.info("CPLUS WebDriver 初始化成功")
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        wait = WebDriverWait(driver, 10)
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
                    logging.error(f"CPLUS {section_name} 嘗試 {attempt+1}/{MAX_RETRIES} 失敗: {str(e)}")
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(1)
            if not success and section_name != 'house':
                raise Exception(f"CPLUS {section_name} 經過 {MAX_RETRIES} 次嘗試失敗")
            elif not success and section_name == 'house':
                logging.warning(f"CPLUS {section_name} 經過 {MAX_RETRIES} 次嘗試失敗，跳過")
        return downloaded_files, house_file_count, house_button_count, driver
    except Exception as e:
        logging.error(f"CPLUS 總錯誤: {str(e)}")
        return downloaded_files, house_file_count, house_button_count, driver
    finally:
        if driver:
            try:
                if 'login_failure' not in logging.getLogger().handlers[0].stream.getvalue():  # 僅登錄成功時登出
                    print("CPLUS: 嘗試登出...", flush=True)
                    logout_menu_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_menu_button)
                    time.sleep(0.2)
                    driver.execute_script("arguments[0].click();", logout_menu_button)
                    print("CPLUS: 登錄按鈕點擊成功", flush=True)
                    try:
                        logout_option = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='menu-list-grow']/div[6]/li")))
                        driver.execute_script("arguments[0].scrollIntoView(true);", logout_option)
                        time.sleep(0.2)
                        driver.execute_script("arguments[0].click();", logout_option)
                        print("CPLUS: Logout 選項點擊成功", flush=True)
                    except TimeoutException:
                        logging.debug("CPLUS: 主 Logout 選項未找到，嘗試備用定位...")
                        try:
                            logout_option = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Logout')]")))
                            driver.execute_script("arguments[0].scrollIntoView(true);", logout_option)
                            time.sleep(0.2)
                            driver.execute_script("arguments[0].click();", logout_option)
                            print("CPLUS: 備用 Logout 選項 1 點擊成功", flush=True)
                        except TimeoutException:
                            logging.warning("CPLUS: 所有 Logout 選項失敗，強制退出")
                time.sleep(0.5)
                try:
                    login_button_pre = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")), 5)
                    logging.info("CPLUS: Logout 成功，返回登錄頁")
                except TimeoutException:
                    logging.warning("CPLUS: Logout 失敗，未返回登錄頁")
            except WebDriverException as e:
                logging.error(f"CPLUS: 登出過程中 WebDriver 崩潰: {str(e)}")
                driver.quit()
            except Exception as logout_error:
                print(f"CPLUS: 登出失敗: {str(logout_error)}", flush=True)
                driver.save_screenshot("logout_failure.png")
                with open("logout_failure.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
            finally:
                driver.quit()
                logging.info("CPLUS WebDriver 關閉")

def barge_login(driver, wait):
    logging.info("Barge: 嘗試打開網站 https://barge.oneport.com/login...")
    driver.get("https://barge.oneport.com/login")
    logging.info(f"Barge: 網站已成功打開，當前 URL: {driver.current_url}")
    time.sleep(0.2)

    logging.info("Barge: 輸入 COMPANY ID...")
    company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Company ID' or contains(@id, 'mat-input-0')]")))
    company_id_field.send_keys("CKL")
    logging.info("Barge: COMPANY ID 輸入完成")
    time.sleep(0.2)

    logging.info("Barge: 輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='User ID' or contains(@id, 'mat-input-1')]")
    user_id_field.send_keys("barge")
    logging.info("Barge: USER ID 輸入完成")
    time.sleep(0.2)

    logging.info("Barge: 輸入 PW...")
    password_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Password' or contains(@id, 'mat-input-2')]")
    password_field.send_keys(os.environ.get('BARGE_PASSWORD', '123456'))
    logging.info("Barge: PW 輸入完成")
    time.sleep(0.2)

    logging.info("Barge: 點擊 LOGIN 按鈕...")
    login_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'LOGIN') or contains(@class, 'mat-raised-button')]")))
    ActionChains(driver).move_to_element(login_button_barge).click().perform()
    logging.info("Barge: LOGIN 按鈕點擊成功")
    time.sleep(0.2)
    
def process_barge_download(driver, wait, initial_files):
    logging.info("Barge: 直接前往 https://barge.oneport.com/downloadReport...")
    driver.get("https://barge.oneport.com/downloadReport")
    time.sleep(0.2)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    logging.info("Barge: downloadReport 頁面加載完成")

    logging.info("Barge: 選擇 Report Type...")
    report_type_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-form-field[.//mat-label[contains(text(), 'Report Type')]]//div[contains(@class, 'mat-select-trigger')]")))
    ActionChains(driver).move_to_element(report_type_trigger).click().perform()
    logging.info("Barge: Report Type 選擇開始")
    time.sleep(0.2)

    logging.info("Barge: 點擊 Container Detail...")
    container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-option//span[contains(text(), 'Container Detail')]")))
    ActionChains(driver).move_to_element(container_detail_option).click().perform()
    logging.info("Barge: Container Detail 點擊成功")
    time.sleep(0.2)

    logging.info("Barge: 點擊 Download...")
    local_initial = initial_files.copy()
    download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Download']]")))
    ActionChains(driver).move_to_element(download_button_barge).click().perform()
    logging.info("Barge: Download 按鈕點擊成功")

    new_files = wait_for_new_file(barge_download_dir, local_initial)
    if new_files:
        logging.info(f"Barge: Container Detail 下載完成，檔案位於: {barge_download_dir}")
        filtered_files = {f for f in new_files if "ContainerDetailReport" in f}
        for file in filtered_files:
            logging.info(f"Barge: 新下載檔案: {file}")
        if not filtered_files:
            logging.warning("Barge: 未下載預期檔案 (ContainerDetailReport*.csv)，記錄頁面狀態...")
            driver.save_screenshot("barge_download_failure.png")
            raise Exception("Barge: Container Detail 未下載預期檔案")
        return filtered_files
    else:
        logging.warning("Barge: Container Detail 未觸發新文件下載，記錄頁面狀態...")
        driver.save_screenshot("barge_download_failure.png")
        raise Exception("Barge: Container Detail 未觸發新文件下載")

def process_barge():
    driver = None
    downloaded_files = set()
    initial_files = set(os.listdir(barge_download_dir))
    try:
        driver = webdriver.Chrome(options=get_chrome_options(barge_download_dir))
        logging.info("Barge WebDriver 初始化成功")
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        wait = WebDriverWait(driver, 10)
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
        if not success:
            logging.error(f"Barge 下載經過 {MAX_RETRIES} 次嘗試失敗")

        return downloaded_files, driver
    except Exception as e:
        logging.error(f"Barge 總錯誤: {str(e)}")
        return downloaded_files, driver
    finally:
        try:
            if driver:
                logging.info("Barge: 點擊工具欄進行登出...")
                try:
                    logout_toolbar_barge = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar_barge)
                    time.sleep(0.2)
                    driver.execute_script("arguments[0].click();", logout_toolbar_barge)
                    logging.info("Barge: 工具欄點擊成功")
                except TimeoutException:
                    logging.debug("Barge: 主工具欄登出按鈕未找到，嘗試備用定位...")
                    raise
                time.sleep(0.2)
                logging.info("Barge: 點擊 Logout 選項...")
                try:
                    logout_span_xpath = "//div[contains(@class, 'mat-menu-panel')]//button//span[contains(text(), 'Logout')]"
                    logout_button_barge = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, logout_span_xpath)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                    time.sleep(0.2)
                    driver.execute_script("arguments[0].click();", logout_button_barge)
                    logging.info("Barge: Logout 選項點擊成功")
                except TimeoutException:
                    logging.debug("Barge: Logout 選項未找到，嘗試備用定位...")
                    backup_logout_xpath = "//button[.//span[contains(text(), 'Logout')]]"
                    logout_button_barge = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, backup_logout_xpath)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                    time.sleep(0.2)
                    driver.execute_script("arguments[0].click();", logout_button_barge)
                    logging.info("Barge: 備用 Logout 選項點擊成功")
                time.sleep(0.5)
        except Exception as e:
            logging.error(f"Barge: 登出失敗: {str(e)}")

def main():
    load_dotenv()
    clear_download_dirs()

    cplus_files = set()
    house_file_count = [0]
    house_button_count = [0]
    barge_files = set()
    cplus_driver = None
    barge_driver = None

    def update_cplus(result):
        files, count, button_count, drv = result
        cplus_files.update(files)
        house_file_count[0] = count
        house_button_count[0] = button_count
        nonlocal cplus_driver
        cplus_driver = drv

    def update_barge(result):
        files, drv = result
        barge_files.update(files)
        nonlocal barge_driver
        barge_driver = drv

    cplus_thread = threading.Thread(target=lambda: update_cplus(process_cplus()))
    barge_thread = threading.Thread(target=lambda: update_barge(process_barge()))

    cplus_thread.start()
    barge_thread.start()

    cplus_thread.join()
    barge_thread.join()

    logging.info("檢查所有下載文件...")
    downloaded_files = [f for f in os.listdir(cplus_download_dir) if f.endswith(('.csv', '.xlsx'))] + [f for f in os.listdir(barge_download_dir) if f.endswith(('.csv', '.xlsx'))]
    logging.info(f"總下載文件: {len(downloaded_files)} 個")
    for file in downloaded_files:
        logging.info(f"找到檔案: {file}")

    required_patterns = {'movement': 'cntrMoveLog', 'onhand': 'data_', 'barge': 'ContainerDetailReport'}
    housekeep_prefixes = ['IE2_', 'DM1C_', 'IA17_', 'GA1_', 'IA5_', 'IA15_']

    has_required = all(any(pattern in f for f in downloaded_files) for pattern in required_patterns.values())
    house_files = [f for f in downloaded_files if any(p in f for p in housekeep_prefixes)]
    house_download_count = len(set(house_files))
    house_ok = (house_download_count >= house_button_count[0] - 1) or (house_button_count[0] == 0)

    if has_required and house_ok:
        logging.info("所有必須文件齊全且 Housekeep 文件數量匹配按鈕數，開始發送郵件...")
        # (郵件發送代碼保持不變)
    else:
        logging.warning(f"文件不齊全: 缺少必須文件 (has_required={has_required}) 或 House文件數量不匹配 (expected={house_button_count[0]}, actual={house_file_count[0]})")

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
