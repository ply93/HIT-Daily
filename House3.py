import os
import time
import shutil
import subprocess
import random
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
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import logging
from dotenv import load_dotenv

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

cplus_download_dir = os.path.abspath("downloads_cplus")
barge_download_dir = os.path.abspath("downloads_barge")
MAX_RETRIES = 3
DOWNLOAD_TIMEOUT = 60  # 延長至 60 秒

def clear_download_dirs():
    for dir_path in [cplus_download_dir, barge_download_dir]:
        if os.path.exists(dir_path):
            shutil.rmtree(dir_path)
        os.makedirs(dir_path)
        logging.info(f"創建下載目錄: {dir_path}")

def setup_environment():
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
    chrome_options.add_argument('--headless=new')  # 改成新 headless 模式
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--disable-popup-blocking')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--no-first-run')
    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36"
    ]
    chrome_options.add_argument(f'--user-agent={random.choice(user_agents)}')
    chrome_options.add_argument(f'--window-size={random.randint(1200, 1920)},{random.randint(800, 1080)}')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
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
        if new_files:
            return new_files
        time.sleep(1)
    return set()

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
    logging.info("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/")
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    logging.info(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}")
    time.sleep(2)
    logging.info("CPLUS: 點擊登錄前按鈕...")
    login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
    ActionChains(driver).move_to_element(login_button_pre).click().perform()
    logging.info("CPLUS: 登錄前按鈕點擊成功")
    time.sleep(2)
    logging.info("CPLUS: 輸入 COMPANY CODE...")
    company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    company_code_field.send_keys("CKL")
    logging.info("CPLUS: COMPANY CODE 輸入完成")
    time.sleep(1)
    logging.info("CPLUS: 輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
    user_id_field.send_keys("KEN")
    logging.info("CPLUS: USER ID 輸入完成")
    time.sleep(1)
    logging.info("CPLUS: 輸入 PASSWORD...")
    password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
    password_field.send_keys(os.environ.get('SITE_PASSWORD'))
    logging.info("CPLUS: PASSWORD 輸入完成")
    time.sleep(1)
    logging.info("CPLUS: 點擊 LOGIN 按鈕...")
    login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
    ActionChains(driver).move_to_element(login_button).click().perform()
    logging.info("CPLUS: LOGIN 按鈕點擊成功")
    time.sleep(2)
    # 加: 檢查 session 是否有效（看用戶菜單元素是否存在）
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))  # 用戶按鈕
        logging.info("CPLUS: 登入 session 有效")
    except TimeoutException:
        logging.error("CPLUS: 登入後 session 失效或 cookie 問題，記錄狀態...")
        driver.save_screenshot("login_session_failure.png")
        with open("login_session_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: 登入 session 失效")

def simulate_user_activity(driver):
    ActionChains(driver).move_by_offset(random.randint(-50, 50), random.randint(-50, 50)).perform()
    driver.execute_script("window.scrollBy(0, 100);")
    time.sleep(random.uniform(1, 3))

def process_cplus_movement(driver, wait, initial_files):
    logging.info("CPLUS: 直接前往 Container Movement Log...")
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
    time.sleep(2)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")))
    logging.info("CPLUS: Container Movement Log 頁面加載完成")

    logging.info("CPLUS: 點擊 Search...")
    local_initial = initial_files.copy()
    for attempt in range(2):
        try:
            search_button = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")))
            ActionChains(driver).move_to_element(search_button).click().perform()
            logging.info("CPLUS: Search 按鈕點擊成功")
            break
        except TimeoutException:
            logging.debug(f"CPLUS: Search 按鈕未找到，嘗試備用定位 {attempt+1}/2...")
            try:
                search_button = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and .//span[contains(text(), 'Search')]]")))
                ActionChains(driver).move_to_element(search_button).click().perform()
                logging.info("CPLUS: 備用 Search 按鈕 1 點擊成功")
                break
            except TimeoutException:
                logging.debug(f"CPLUS: 備用 Search 按鈕 1 失敗，嘗試備用定位 2 (嘗試 {attempt+1}/2)...")
                try:
                    search_button = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]")))
                    ActionChains(driver).move_to_element(search_button).click().perform()
                    logging.info("CPLUS: 備用 Search 按鈕 2 點擊成功")
                    break
                except TimeoutException:
                    logging.debug(f"CPLUS: 備用 Search 按鈕 2 失敗 (嘗試 {attempt+1}/2)")
                    driver.save_screenshot("movement_search_failure.png")
                    with open("movement_search_failure.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
    else:
        raise Exception("CPLUS: Container Movement Log Search 按鈕點擊失敗")

    simulate_user_activity(driver)

    logging.info("CPLUS: 點擊 Download...")
    for attempt in range(2):
        try:
            download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button")))
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
    time.sleep(1)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    # 加: 檢查 JS 執行或相容問題
    try:
        # 檢查 document.readyState 是否 complete（JS 載入完成）
        js_state = driver.execute_script("return document.readyState;")
        if js_state != "complete":
            logging.warning("CPLUS OnHand: JS 未完全執行，狀態: {js_state}，嘗試等待...")
            time.sleep(10) # 加延遲，等 JS 跑
            # 再檢查
            js_state = driver.execute_script("return document.readyState;")
            if js_state != "complete":
                raise Exception("CPLUS OnHand: JS 執行失敗，狀態: {js_state}")
        # 檢查 noscript 是否 visible（如果 visible，JS 未跑）
        try:
            wait.until_not(EC.visibility_of_element_located((By.TAG_NAME, "noscript"))) # 如果 noscript 可見，JS 未跑
        except TimeoutException:
            logging.error("CPLUS OnHand: noscript 可見，JS 執行或相容問題，記錄狀態...")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S") # 修正：移除多餘 datetime.datetime
            driver.save_screenshot(f"onhand_js_failure_{timestamp}.png")
            with open(f"onhand_js_failure_{timestamp}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            # 試 refresh 解決
            logging.warning("CPLUS OnHand: 嘗試刷新頁面解決 JS 問題...")
            driver.refresh()
            time.sleep(5)
            try:
                wait.until_not(EC.visibility_of_element_located((By.TAG_NAME, "noscript")))
            except TimeoutException:
                raise Exception("CPLUS OnHand: JS 執行或相容問題，noscript 仍可見")
        time.sleep(5)  # 優化：從 10 秒縮短到 5 秒，等 JS 完全渲染
        # 改: 檢查渲染元素是否存在（用更準 XPath match Search），縮短等待時間，並加備用 locator
        try:
            extended_wait = WebDriverWait(driver, 15)
            search_element_locators = [
                (By.XPATH, "//button//span[contains(text(), 'Search')]"),  # 原有
                (By.CSS_SELECTOR, "button.MuiButton-containedPrimary span.MuiButton-label"),  # 備用 CSS，基於 Material-UI
                (By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and .//span[contains(text(), 'Search')]]")  # 另一備用
            ]
            found = False
            for locator in search_element_locators:
                try:
                    extended_wait.until(EC.presence_of_element_located(locator))
                    logging.info(f"CPLUS OnHand: 渲染元素存在（使用 locator: {locator}），JS 執行正常")
                    found = True
                    break
                except TimeoutException:
                    logging.debug(f"CPLUS OnHand: 備用 locator {locator} 未找到，試下一個...")
            if not found:
                raise TimeoutException("All locators failed")  # 觸發下一個 except
        except TimeoutException:
            # 新加：如果超時，嘗試刷新頁面再檢查
            logging.warning("CPLUS OnHand: 渲染元素未出現，嘗試刷新頁面...")
            driver.refresh()
            time.sleep(10)
            try:
                extended_wait.until(EC.presence_of_element_located((By.XPATH, "//button//span[contains(text(), 'Search')]")))
                logging.info("CPLUS OnHand: 刷新後渲染元素存在，JS 執行正常")
            except TimeoutException:
                # 新加 debug 部分：儲存 screenshot 同 page_source
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # 修正：移除多餘 datetime.datetime
                driver.save_screenshot(f"onhand_render_failure_{timestamp}.png")
                with open(f"onhand_render_failure_{timestamp}.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                logging.error("CPLUS OnHand: 渲染元素未出現，已儲存 screenshot 同 page_source 作 debug")
                raise Exception("CPLUS OnHand: JS 執行問題，無渲染元素")
    except Exception as e:
        logging.error(f"CPLUS OnHand: JS 檢查失敗: {str(e)}")
        raise # 繼續 raise，讓 retry
    logging.info("CPLUS: OnHandContainerList 頁面加載完成")
    logging.info("CPLUS: 點擊 Search...")
    local_initial = initial_files.copy()
    try:
        search_button_onhand = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")))
        time.sleep(0.5)
        ActionChains(driver).move_to_element(search_button_onhand).click().perform()
        logging.info("CPLUS: Search 按鈕點擊成功")
    except TimeoutException:
        logging.debug("CPLUS: Search 按鈕未找到，嘗試備用定位...")
        try:
            search_button_onhand = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and .//span[contains(text(), 'Search')]]")))
            time.sleep(0.5)
            ActionChains(driver).move_to_element(search_button_onhand).click().perform()
            logging.info("CPLUS: 備用 Search 按鈕 1 點擊成功")
        except TimeoutException:
            logging.debug("CPLUS: 備用 Search 按鈕 1 失敗，嘗試第三備用定位...")
            try:
                search_button_onhand = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.MuiButton-contained span.MuiButton-label")))
                time.sleep(0.5)
                ActionChains(driver).move_to_element(search_button_onhand).click().perform()
                logging.info("CPLUS: 第三備用 Search 按鈕點擊成功")
            except TimeoutException:
                logging.error("CPLUS: 所有 Search 按鈕定位失敗，記錄頁面狀態...")
                driver.save_screenshot("onhand_search_failure.png")
                with open("onhand_search_failure.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                raise Exception("CPLUS: OnHandContainerList Search 按鈕點擊失敗")
    time.sleep(0.5)
    simulate_user_activity(driver)
    logging.info("CPLUS: 點擊 Export...")
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button")))
    ActionChains(driver).move_to_element(export_button).click().perform()
    logging.info("CPLUS: Export 按鈕點擊成功")
    time.sleep(0.5)
    logging.info("CPLUS: 點擊 Export as CSV...")
    export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']")))
    ActionChains(driver).move_to_element(export_csv_button).click().perform()
    logging.info("CPLUS: Export as CSV 按鈕點擊成功")
    time.sleep(0.5)
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
    clean_house_files(cplus_download_dir)  # 新加：清理舊 House 檔案
    logging.info("CPLUS: 前往 Housekeeping Reports 頁面...")
    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    logging.info("CPLUS: Housekeeping Reports 頁面加載完成")
    logging.info("CPLUS: 等待表格加載...")
    try:
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
        if len(rows) == 0 or all(not row.text.strip() for row in rows):
            logging.debug("表格數據空或無效，刷新頁面...")
            driver.refresh()
            wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
            rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
            if len(rows) < 6:
                logging.warning("刷新後表格數據仍不足，記錄頁面狀態...")
                driver.save_screenshot("house_load_failure.png")
                with open("house_load_failure.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                raise Exception("CPLUS: Housekeeping Reports 表格數據不足")
        logging.info("CPLUS: 表格加載完成")
    except TimeoutException:
        logging.warning("CPLUS: 表格未加載，嘗試刷新頁面...")
        driver.refresh()
        wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
        logging.info("CPLUS: 表格加載完成 (after refresh)")
    time.sleep(5)  # 額外等待 JS 渲染按鈕
    logging.info("CPLUS: 等待 Excel 按鈕出現...")
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)]")))
        logging.info("CPLUS: Excel 按鈕已出現")
    except TimeoutException:
        logging.warning("CPLUS: Excel 按鈕未出現，記錄狀態...")
        driver.save_screenshot("house_button_wait_failure.png")
        with open("house_button_wait_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
    logging.info("CPLUS: 定位並點擊所有 Excel 下載按鈕...")
    local_initial = initial_files.copy()
    new_files = set()
    report_files = {} # 儲存報告名稱與檔案名稱的映射
    excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)]")
    button_count = len(excel_buttons)
    logging.info(f"CPLUS: 找到 {button_count} 個 Excel 下載按鈕")
    if button_count == 0:
        logging.debug("CPLUS: 未找到 Excel 按鈕，嘗試備用定位...")
        excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]//svg[@viewBox='0 0 24 24']//path[@fill='#036e11']")
        button_count = len(excel_buttons)
        logging.info(f"CPLUS: 備用定位找到 {button_count} 個 Excel 下載按鈕")
    # 每個按鈕前清視窗
    handle_popup(driver, wait)
    for idx in range(button_count):
        success = False
        for retry in range(3):  # 修改: 加重試 3 次每個按鈕
            try:
                button_xpath = f"(//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)])[{idx+1}]"
                button = wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
                try:
                    report_name = driver.find_element(By.XPATH, f"//table[contains(@class, 'MuiTable-root')]//tbody//tr[{idx+1}]//td[3]").text
                    logging.info(f"CPLUS: 準備點擊第 {idx+1} 個 Excel 按鈕，報告名稱: {report_name}")
                except:
                    logging.debug(f"CPLUS: 無法獲取第 {idx+1} 個按鈕的報告名稱")
                # 修改: 用 JS 點擊避 intercept
                driver.execute_script("arguments[0].click();", button)
                logging.info(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕 JavaScript 點擊成功")
                time.sleep(0.5)  # 加小延遲等待彈出
                handle_popup(driver, wait)
                temp_new = wait_for_new_file(cplus_download_dir, local_initial)
                if temp_new:
                    file_name = temp_new.pop()
                    logging.info(f"CPLUS: 第 {idx+1} 個按鈕下載新文件: {file_name}")
                    local_initial.add(file_name)
                    new_files.add(file_name)
                    report_files[report_name] = file_name
                    success = True
                    break
                else:
                    logging.warning(f"CPLUS: 第 {idx+1} 個按鈕未觸發新文件下載 (重試 {retry+1}/3)")
                    time.sleep(1)
            except Exception as e:
                logging.error(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕點擊失敗 (重試 {retry+1}/3): {str(e)}")
                handle_popup(driver, wait)  # 失敗時再清視窗
                time.sleep(1)
                if retry == 2:  # 最後一次記錄 debug
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    driver.save_screenshot(f"house_button_failure_{idx+1}_{timestamp}.png")
                    with open(f"house_button_failure_{idx+1}_{timestamp}.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
        if not success:
            logging.warning(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕經過 3 次重試失敗")
    if new_files:
        logging.info(f"CPLUS: Housekeeping Reports 下載完成，共 {len(new_files)} 個文件，預期 {button_count} 個")
        if len(new_files) != button_count:
            logging.error(f"CPLUS: 下載數 {len(new_files)} 不等於按鈕數 {button_count}，可能漏下載，記錄狀態...")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            driver.save_screenshot(f"house_download_mismatch_{timestamp}.png")
            with open(f"house_download_mismatch_{timestamp}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            raise Exception("CPLUS: Housekeeping Reports 下載數不匹配")
        return new_files, len(new_files), button_count, report_files
    else:
        logging.warning("CPLUS: Housekeeping Reports 未下載任何文件，記錄頁面狀態...")
        driver.save_screenshot("house_download_failure.png")
        with open("house_download_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: Housekeeping Reports 未下載任何文件")

def process_cplus():
    driver = None
    downloaded_files = set()
    initial_files = set(os.listdir(cplus_download_dir))
    house_file_count = 0
    house_button_count = 0
    house_report_files = {}
    try:
        driver = webdriver.Chrome(options=get_chrome_options(cplus_download_dir))
        logging.info("CPLUS WebDriver 初始化成功")
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        wait = WebDriverWait(driver, 15)
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
                    # 加: 在每個 attempt 前檢查 session
                    try:
                        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]"))) # 用戶按鈕元素
                        logging.info(f"CPLUS {section_name}: Session 有效，繼續")
                    except TimeoutException:
                        logging.warning(f"CPLUS {section_name}: Session 失效或 cookie 問題，重新登入...")
                        cplus_login(driver, wait) # 自動重新登入
                        # 加記錄，幫助 debug
                        driver.save_screenshot(f"session_failure_{section_name}_attempt{attempt+1}.png")
                        with open(f"session_failure_{section_name}_attempt{attempt+1}.html", "w", encoding="utf-8") as f:
                            f.write(driver.page_source)
                    # 修改: 加延遲同刷新，避免載入崩潰
                    if section_name == 'movement':
                        time.sleep(5)  # 加延遲讓頁面穩定
                    new_files = section_func(driver, wait, initial_files) if section_name != 'house' else section_func(driver, wait, initial_files)
                    if section_name == 'house':
                        new_files, house_file_count, house_button_count, house_report_files = new_files  # unpack
                    downloaded_files.update(new_files)
                    initial_files.update(new_files)
                    success = True
                    break
                except Exception as e:
                    logging.error(f"CPLUS {section_name} 嘗試 {attempt+1}/{MAX_RETRIES} 失敗: {str(e)}")
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(5)
                        # 修改: 加刷新頁面或重新導航，避免內部崩潰殘留
                        try:
                            driver.refresh()
                        except:
                            pass
            if not success:
                logging.error(f"CPLUS {section_name} 經過 {MAX_RETRIES} 次嘗試失敗")
        return downloaded_files, house_file_count, house_button_count, driver, house_report_files
    except Exception as e:
        logging.error(f"CPLUS 總錯誤: {str(e)}")
        return downloaded_files, house_file_count, house_button_count, driver, {}
    finally:
        try:
            if driver:
                logging.info("CPLUS: 嘗試登出...")
                logout_menu_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
                logout_menu_button.click()
                logging.info("CPLUS: 用戶菜單點擊成功")
                logout_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Logout')]")))
                logout_option.click()
                logging.info("CPLUS: Logout 選項點擊成功")
                time.sleep(1)  # 等待視窗出現
                close_success = False
                for retry in range(3):
                    try:
                        close_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="logout"]/div[3]/div/div[3]/button/span[1]')))
                        driver.execute_script("arguments[0].click();", close_button)
                        logging.info("CPLUS: Logout 後 CLOSE 按鈕 JavaScript 點擊成功")
                        close_success = True
                        break
                    except Exception as ce:
                        logging.warning(f"CPLUS: CLOSE 按鈕點擊失敗 (重試 {retry+1}/3): {str(ce)}")
                        handle_popup(driver, wait)
                        time.sleep(0.5)
                if not close_success:
                    logging.error("CPLUS: CLOSE 按鈕經過 3 次重試失敗，記錄狀態...")
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    driver.save_screenshot(f"logout_close_failure_{timestamp}.png")
                    with open(f"logout_close_failure_{timestamp}.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
        except Exception as e:
            logging.error(f"CPLUS: 登出失敗: {str(e)}")
            
def barge_login(driver, wait):
    logging.info("Barge: 嘗試打開網站 https://barge.oneport.com/login...")
    driver.get("https://barge.oneport.com/login")
    logging.info(f"Barge: 網站已成功打開，當前 URL: {driver.current_url}")
    time.sleep(3)

    logging.info("Barge: 輸入 COMPANY ID...")
    company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Company ID' or contains(@id, 'mat-input-0')]")))
    company_id_field.send_keys("CKL")
    logging.info("Barge: COMPANY ID 輸入完成")
    time.sleep(1)

    logging.info("Barge: 輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='User ID' or contains(@id, 'mat-input-1')]")
    user_id_field.send_keys("barge")
    logging.info("Barge: USER ID 輸入完成")
    time.sleep(1)

    logging.info("Barge: 輸入 PW...")
    password_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Password' or contains(@id, 'mat-input-2')]")
    password_field.send_keys(os.environ.get('BARGE_PASSWORD', '123456'))
    logging.info("Barge: PW 輸入完成")
    time.sleep(1)

    logging.info("Barge: 點擊 LOGIN 按鈕...")
    login_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'LOGIN') or contains(@class, 'mat-raised-button')]")))
    ActionChains(driver).move_to_element(login_button_barge).click().perform()
    logging.info("Barge: LOGIN 按鈕點擊成功")
    time.sleep(3)

def process_barge_download(driver, wait, initial_files):
    logging.info("Barge: 直接前往 https://barge.oneport.com/downloadReport...")
    driver.get("https://barge.oneport.com/downloadReport")
    time.sleep(3)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    logging.info("Barge: downloadReport 頁面加載完成")

    logging.info("Barge: 選擇 Report Type...")
    report_type_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-form-field[.//mat-label[contains(text(), 'Report Type')]]//div[contains(@class, 'mat-select-trigger')]")))
    ActionChains(driver).move_to_element(report_type_trigger).click().perform()
    logging.info("Barge: Report Type 選擇開始")
    time.sleep(2)

    logging.info("Barge: 點擊 Container Detail...")
    container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-option//span[contains(text(), 'Container Detail')]")))
    ActionChains(driver).move_to_element(container_detail_option).click().perform()
    logging.info("Barge: Container Detail 點擊成功")
    time.sleep(2)

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
        wait = WebDriverWait(driver, 15)

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
                    time.sleep(5)
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
                    logout_toolbar_barge = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar_barge)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", logout_toolbar_barge)
                    logging.info("Barge: 工具欄點擊成功")
                except TimeoutException:
                    logging.debug("Barge: 主工具欄登出按鈕未找到，嘗試備用定位...")
                    raise

                time.sleep(2)

                logging.info("Barge: 點擊 Logout 選項...")
                try:
                    logout_span_xpath = "//div[contains(@class, 'mat-menu-panel')]//button//span[contains(text(), 'Logout')]"
                    logout_button_barge = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, logout_span_xpath)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", logout_button_barge)
                    logging.info("Barge: Logout 選項點擊成功")
                except TimeoutException:
                    logging.debug("Barge: Logout 選項未找到，嘗試備用定位...")
                    backup_logout_xpath = "//button[.//span[contains(text(), 'Logout')]]"
                    logout_button_barge = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, backup_logout_xpath)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", logout_button_barge)
                    logging.info("Barge: 備用 Logout 選項點擊成功")

                time.sleep(5)

        except Exception as e:
            logging.error(f"Barge: 登出失敗: {str(e)}")

def main():
    load_dotenv()
    clear_download_dirs()
    cplus_files = set()
    house_file_count = 0
    house_button_count = 0
    barge_files = set()
    cplus_driver = None
    house_report_files = {}
    # Process CPLUS
    cplus_files, house_file_count, house_button_count, cplus_driver, house_report_files = process_cplus()
    if cplus_driver:
        cplus_driver.quit()
        logging.info("CPLUS WebDriver 關閉")
    # Process Barge
    barge_files, barge_driver = process_barge()
    if barge_driver:
        barge_driver.quit()
        logging.info("Barge WebDriver 關閉")
    # Check all downloaded files
    logging.info("檢查所有下載文件...")
    downloaded_files = [f for f in os.listdir(cplus_download_dir) if f.endswith(('.csv', '.xlsx'))] + [f for f in os.listdir(barge_download_dir) if f.endswith(('.csv', '.xlsx'))]
    logging.info(f"總下載文件: {len(downloaded_files)} 個")
    for file in downloaded_files:
        logging.info(f"找到檔案: {file}")
    required_patterns = {'movement': 'cntrMoveLog', 'onhand': 'data_', 'barge': 'ContainerDetailReport'}
    housekeep_prefixes = ['IE2_', 'DM1C_', 'IA17_', 'GA1_', 'IA5_', 'IA15_', 'INV-114_']  # 修改：添加 'INV-114_'
    has_required = all(any(pattern in f for f in downloaded_files) for pattern in required_patterns.values())
    # 新加：收集獨特 House 檔案，按前綴選最新
    house_files_dict = {}
    for file in os.listdir(cplus_download_dir):
        for prefix in housekeep_prefixes:
            if file.startswith(prefix) and file.endswith('.csv'):
                file_path = os.path.join(cplus_download_dir, file)
                mod_time = os.path.getmtime(file_path)
                if prefix not in house_files_dict or mod_time > house_files_dict[prefix]['mod_time']:
                    house_files_dict[prefix] = {'file': file, 'mod_time': mod_time}
    house_unique_files = [info['file'] for info in house_files_dict.values()]
    house_download_count = len(house_unique_files)
    house_ok = house_download_count >= 6  # 修改：改為 >= 6，忽略多餘
    if has_required and house_ok:
        logging.info("所有必須文件齊全，開始發送郵件...")
        try:
            smtp_server = os.environ.get('SMTP_SERVER', 'smtp.zoho.com')
            smtp_port = int(os.environ.get('SMTP_PORT', 587))
            sender_email = os.environ['ZOHO_EMAIL']
            sender_password = os.environ['ZOHO_PASSWORD']
            receiver_emails = os.environ.get('RECEIVER_EMAILS', 'ckeqc@ckline.com.hk').split(',')
            cc_emails = os.environ.get('CC_EMAILS', '').split(',') if os.environ.get('CC_EMAILS') else []
            dry_run = os.environ.get('DRY_RUN', 'False').lower() == 'true'
            if dry_run:
                logging.info("Dry run 模式：只打印郵件內容，不發送。")
            # 動態生成表格內容（改用獨特檔案）
            body_html = f"""
            <html><body><p>Attached are the daily reports downloaded from CPLUS and Barge. Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <table border="1" style="border-collapse: collapse; width: 100%;"><thead><tr><th>Category</th><th>Report</th><th>File Names</th><th>Status</th></tr></thead><tbody>
            <tr><td rowspan="{2 + len(house_report_files)}">CPLUS</td><td>Container Movement</td><td>{', '.join([f for f in downloaded_files if 'cntrMoveLog' in f]) or 'N/A'}</td><td>{'✓' if any('cntrMoveLog' in f for f in downloaded_files) else '-'}</td></tr>
            <tr><td>OnHandContainerList</td><td>{', '.join([f for f in downloaded_files if 'data_' in f]) or 'N/A'}</td><td>{'✓' if any('data_' in f for f in downloaded_files) else '-'}</td></tr>
            """
            for report_name, file_name in house_report_files.items():
                if file_name in house_unique_files:  # 只列獨特
                    status = '✓' if file_name in downloaded_files else '-'
                    body_html += f"<tr><td>{report_name}</td><td>{file_name}</td><td>{status}</td></tr>\n"
            body_html += f"""
            <tr><td rowspan="1">BARGE</td><td>Container Detail</td><td>{', '.join([f for f in downloaded_files if 'ContainerDetailReport' in f]) or 'N/A'}</td><td>{'✓' if any('ContainerDetailReport' in f for f in downloaded_files) else '-'}</td></tr>
            <tr><td colspan="2"><strong>TOTAL</strong></td><td><strong>{len(downloaded_files)} files attached</strong></td><td><strong>{len(downloaded_files)}</strong></td></tr>
            </tbody></table></body></html>
            """
            msg = MIMEMultipart('alternative')
            msg['From'] = sender_email
            msg['To'] = ', '.join(receiver_emails)
            if cc_emails:
                msg['Cc'] = ', '.join(cc_emails)
            msg['Subject'] = f"HIT DAILY {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            msg.attach(MIMEText(body_html, 'html'))
            plain_text = body_html.replace('<br>', '\n').replace('<table>', '').replace('</table>', '').replace('<tr>', '\n').replace('<td>', ' | ').replace('</td>', '').replace('<th>', ' | ').replace('</th>', '').strip()
            msg.attach(MIMEText(plain_text, 'plain'))
            # 修改：附件只加獨特 House + 其他
            attachments = house_unique_files[:]  # House 獨特
            # 加 OnHand
            for file in os.listdir(cplus_download_dir):
                if file.startswith('data_') and file.endswith('.csv'):
                    attachments.append(file)
                    break
            # 加 Movement
            for file in os.listdir(cplus_download_dir):
                if 'cntrMoveLog' in file and file.endswith('.xlsx'):
                    attachments.append(file)
                    break
            # 加 Barge
            for file in os.listdir(barge_download_dir):
                if file.startswith('ContainerDetailReport') and file.endswith('.csv'):
                    attachments.append(file)
                    break
            for file in attachments:
                if file in os.listdir(cplus_download_dir):
                    file_path = os.path.join(cplus_download_dir, file)
                else:
                    file_path = os.path.join(barge_download_dir, file)
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
        logging.warning(f"文件不齊全: 缺少必須文件 (has_required={has_required}) 或 House文件不足 (download={house_download_count}, button={house_button_count})")
    logging.info("腳本完成")

if __name__ == "__main__":
    setup_environment()
    main()
