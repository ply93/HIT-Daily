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
from selenium.common.exceptions import TimeoutException, JavascriptException, ElementClickInterceptedException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import logging
from dotenv import load_dotenv

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

cplus_download_dir = os.path.abspath("downloads_cplus")
barge_download_dir = os.path.abspath("downloads_barge")
MAX_RETRIES = 3
DOWNLOAD_TIMEOUT = 30  # 延長至 60 秒

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
            raise Exception("Chromium 未安裝，請檢查 GitHub Actions YML 安裝步驟")
        else:
            logging.info("Chromium 及 ChromeDriver 已存在，跳過安裝")
        
        result = subprocess.run(['pip', 'show', 'selenium'], capture_output=True, text=True)
        if "selenium" not in result.stdout:
            raise Exception("Selenium 未安裝，請檢查 GitHub Actions YML pip 步驟")
        result = subprocess.run(['pip', 'show', 'webdriver-manager'], capture_output=True, text=True)
        if "webdriver-manager" not in result.stdout:
            raise Exception("WebDriver Manager 未安裝，請檢查 GitHub Actions YML pip 步驟")
        logging.info("Selenium 及 WebDriver Manager 已存在，跳過安裝")
    except Exception as e:
        logging.error(f"環境檢查失敗: {e}")
        raise

# 完整 SUB CODE: 修改 get_chrome_options 函數，調整 random.randint(800, 1440) 範圍（替換原 get_chrome_options 全部內容）
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
    chrome_options.add_argument(f'--window-size={random.randint(1440, 2560)},{random.randint(1440, 2560)}')
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

def wait_for_new_file(download_dir, initial_files, timeout=20, prefixes=None):
    start_time = time.time()
    while time.time() - start_time < timeout:
        # 先檢查有冇下載中檔案 (.crdownload)
        downloading_files = [f for f in os.listdir(download_dir) if f.endswith('.crdownload')]
        if downloading_files:
            logging.debug(f"檢測到下載中檔案: {downloading_files}，繼續等待...")
            time.sleep(1)
            continue  # 如果有下載中，跳過新檔案檢查，等完成
        
        # 檢查完成檔案
        current_files = set(f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx')))
        new_files = current_files - initial_files
        
        if new_files:
            # 確保每個新檔案大小穩定（下載完成）
            stable_files = []
            for file in new_files:
                file_path = os.path.join(download_dir, file)
                size1 = os.path.getsize(file_path)
                time.sleep(1)  # 等 1 秒再 check 大小
                size2 = os.path.getsize(file_path)
                if size1 == size2 and size1 > 0:  # 大小穩定且非空
                    stable_files.append(file)
                else:
                    logging.debug(f"檔案 {file} 大小不穩定 ({size1} -> {size2})，繼續等待...")
            
            if stable_files:
                if prefixes:
                    filtered_new = [f for f in stable_files if any(f.startswith(p) for p in prefixes)]
                    if filtered_new:
                        logging.info(f"下載完成並穩定: {filtered_new}")
                        return set(filtered_new)
                else:
                    logging.info(f"下載完成並穩定: {stable_files}")
                    return set(stable_files)
        
        time.sleep(1)  # 每秒檢查一次
    
    logging.warning(f"等待 {timeout} 秒後仍無新穩定檔案，返回空集")
    return set()

# 完整 sub code: 修改 handle_popup 函數，加記錄彈出內容（替換原 handle_popup）
def handle_popup(driver, wait):
    try:
        popup = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'System Error') or contains(@class, 'MuiDialog-container') or contains(@class, 'MuiDialog') and not(@aria-label='menu')]"))
        )
        popup_text = popup.text  # 記錄彈出內容
        logging.info(f"檢測到彈出視窗，內容: {popup_text}")
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
    except Exception as e:
        logging.error(f"處理彈出視窗意外錯誤: {str(e)}")

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
    time.sleep(1)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")))
    logging.info("CPLUS: Container Movement Log 頁面加載完成")

    logging.info("CPLUS: 點擊 Search...")
    local_initial = initial_files.copy()
    for attempt in range(2):
        try:
            search_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")))
            ActionChains(driver).move_to_element(search_button).click().perform()
            logging.info("CPLUS: Search 按鈕點擊成功")
            break
        except TimeoutException:
            logging.debug(f"CPLUS: Search 按鈕未找到，嘗試備用定位 {attempt+1}/2...")
            try:
                search_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and .//span[contains(text(), 'Search')]]")))
                ActionChains(driver).move_to_element(search_button).click().perform()
                logging.info("CPLUS: 備用 Search 按鈕 1 點擊成功")
                break
            except TimeoutException:
                logging.debug(f"CPLUS: 備用 Search 按鈕 1 失敗，嘗試備用定位 2 (嘗試 {attempt+1}/2)...")
                try:
                    search_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]")))
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
            time.sleep(5)
            # 再檢查
            js_state = driver.execute_script("return document.readyState;")
            if js_state != "complete":
                raise Exception("CPLUS OnHand: JS 執行失敗，狀態: {js_state}")

        try:
            wait.until_not(EC.visibility_of_element_located((By.TAG_NAME, "noscript")))
        except TimeoutException:
            logging.error("CPLUS OnHand: noscript 可見，JS 執行或相容問題，記錄狀態...")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S") # 修正：移除多餘 datetime.datetime
            driver.save_screenshot(f"onhand_js_failure_{timestamp}.png")
            with open(f"onhand_js_failure_{timestamp}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            # 試 refresh 解決
            logging.warning("CPLUS OnHand: 嘗試刷新頁面解決 JS 問題...")
            driver.refresh()
            time.sleep(3)
            try:
                wait.until_not(EC.visibility_of_element_located((By.TAG_NAME, "noscript")))
            except TimeoutException:
                raise Exception("CPLUS OnHand: JS 執行或相容問題，noscript 仍可見")
        time.sleep(5)
        try:
            extended_wait = WebDriverWait(driver, 10)
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
            time.sleep(5)
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
        search_button_onhand = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")))
        time.sleep(0.5)
        ActionChains(driver).move_to_element(search_button_onhand).click().perform()
        logging.info("CPLUS: Search 按鈕點擊成功")
    except TimeoutException:
        logging.debug("CPLUS: Search 按鈕未找到，嘗試備用定位...")
        try:
            search_button_onhand = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and .//span[contains(text(), 'Search')]]")))
            time.sleep(0.5)
            ActionChains(driver).move_to_element(search_button_onhand).click().perform()
            logging.info("CPLUS: 備用 Search 按鈕 1 點擊成功")
        except TimeoutException:
            logging.debug("CPLUS: 備用 Search 按鈕 1 失敗，嘗試第三備用定位...")
            try:
                search_button_onhand = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.MuiButton-contained span.MuiButton-label")))
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
    logging.info("CPLUS: 前往 Housekeeping Reports 頁面...")
    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    logging.info("CPLUS: Housekeeping Reports 頁面加載完成")
    
    def wait_for_table_and_buttons(load_retry_max=3):
        """共用函數：等待表格加載同 Excel 按鈕出現"""
        success_load = False
        for load_retry in range(load_retry_max):
            try:
                # 等待表格行出現，並檢查至少 6 行（預期報告數）
                rows = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
                if len(rows) < 6 or all(not row.text.strip() for row in rows):
                    logging.debug(f"表格數據不足或空 (現在 {len(rows)} 行 < 預期 6)，刷新頁面重試...")
                    driver.refresh()
                    # 再等待
                    rows = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
                    if len(rows) < 6:
                        logging.warning("刷新後表格數據仍不足，記錄頁面狀態...")
                        driver.save_screenshot("house_load_failure.png")
                        with open("house_load_failure.html", "w", encoding="utf-8") as f:
                            f.write(driver.page_source)
                        break  # 不 raise，繼續
                logging.info("CPLUS: 表格加載完成")
                success_load = True
                break
            except TimeoutException:
                logging.warning(f"CPLUS: 表格未加載 (重試 {load_retry+1}/{load_retry_max})，嘗試刷新頁面...")
                driver.refresh()
        if not success_load:
            logging.error("CPLUS: Housekeeping Reports 表格加載失敗3次，繼續其他邏輯...")
        
        logging.info("CPLUS: 等待 Excel 按鈕出現...")
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)]")))
            logging.info("CPLUS: Excel 按鈕已出現")
        except TimeoutException:
            logging.warning("CPLUS: Excel 按鈕未出現，記錄狀態...")
            driver.save_screenshot("house_button_wait_failure.png")
            with open("house_button_wait_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
    
    # 初始等待
    wait_for_table_and_buttons()
    
    logging.info("CPLUS: 定位並點擊所有 Excel 下載按鈕...")
    local_initial = initial_files.copy()
    new_files = set()
    report_files = {}  # 儲存報告名稱與 {'file': file_name, 'mod_time': mod_time} 的映射
    # 使用與等待一致的XPATH定位按鈕
    button_locator = "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)]"
    buttons = driver.find_elements(By.XPATH, button_locator)
    total_buttons = len(buttons)
    logging.info(f"CPLUS: 搵到 {total_buttons} 個 Excel 下載按鈕 (使用穩定XPATH)")
    if total_buttons == 0:
        logging.debug("CPLUS: 未找到 Excel 按鈕，嘗試備用CSS...")
        buttons = driver.find_elements(By.CSS_SELECTOR, "button[title='Excel']")  # 保留原備用
        total_buttons = len(buttons)
        logging.info(f"CPLUS: 備用CSS找到 {total_buttons} 個 Excel 下載按鈕")
    if total_buttons == 0:
        logging.warning("CPLUS: 定位失敗，找到0個按鈕，記錄debug資訊...")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        driver.save_screenshot(f"house_locator_failure_{timestamp}.png")
        with open(f"house_locator_failure_{timestamp}.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
    # 每個按鈕前清視窗
    handle_popup(driver, wait)
    housekeep_prefixes = ['IE2_', 'DM1C_', 'IA17_', 'GA1_', 'IA5_', 'IA15_', 'INV-114_']  # 用於過濾
    for i in range(total_buttons):
        success = False
        for retry in range(3):  # 加重試 3 次每個按鈕
            try:
                # 【重要】每次 loop 都重新搵一次所有按鈕，避免元素失效
                current_buttons = driver.find_elements(By.XPATH, button_locator)
                if len(current_buttons) < total_buttons:
                    logging.warning(f"CPLUS: 按鈕數不足 (現在 {len(current_buttons)} < 預期 {total_buttons})，等待再試...")
                    wait_for_table_and_buttons(load_retry_max=1)  # 快速重試一次等待
                    continue  # 跳過本次 retry
                btn = current_buttons[i]
                # 捲動到該按鈕位置，確保佢喺畫面內
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                time.sleep(1)  # 等待滾動完成，避免點擊失敗
                try:
                    report_name = driver.find_element(By.XPATH, f"//table[contains(@class, 'MuiTable-root')]//tbody//tr[{i+1}]//td[3]").text
                    logging.info(f"CPLUS: 準備點擊第 {i+1} 個 Excel 按鈕，報告名稱: {report_name}")
                except:
                    logging.debug(f"CPLUS: 無法獲取第 {i+1} 個按鈕的報告名稱")
                    report_name = f"Unknown Report {i+1}"  # 後備名稱，避免 key error
                # 執行點擊
                driver.execute_script("arguments[0].click();", btn)
                logging.info(f"CPLUS: 第 {i+1} 個 Excel 按鈕點擊成功")
                time.sleep(0.5)  # 加小延遲等待彈出
                # 等待下載或彈窗處理 (保留你原本處理彈窗嘅 code)
                handle_popup(driver, wait)
                temp_new = wait_for_new_file(cplus_download_dir, local_initial, timeout=20, prefixes=housekeep_prefixes)  # 20s
                if temp_new:
                    file_name = temp_new.pop()
                    logging.info(f"CPLUS: 第 {i+1} 個按鈕下載新文件: {file_name}")
                    local_initial.add(file_name)
                    new_files.add(file_name)
                    file_path = os.path.join(cplus_download_dir, file_name)
                    mod_time = os.path.getmtime(file_path)
                    # 如果報告已存在，選最新，並優先無 (1) 的
                    if report_name in report_files:
                        old_file = report_files[report_name]['file']
                        old_mod = report_files[report_name]['mod_time']
                        if mod_time > old_mod or (' (' not in file_name and ' (' in old_file):
                            report_files[report_name] = {'file': file_name, 'mod_time': mod_time}
                    else:
                        report_files[report_name] = {'file': file_name, 'mod_time': mod_time}
                    success = True
                    # 加: 成功後刷新頁面，重置狀態
                    logging.info(f"CPLUS: 第 {i+1} 個下載成功，刷新頁面重置狀態...")
                    driver.refresh()
                    # 新加：刷新後重新等待表格同按鈕，確保下一個 loop 穩定（無固定 sleep）
                    wait_for_table_and_buttons(load_retry_max=2)  # 減到2次，避免太長
                    break
                else:
                    logging.warning(f"CPLUS: 第 {i+1} 個按鈕未觸發新文件下載 (重試 {retry+1}/3)")
                    time.sleep(1)
                time.sleep(3)  # 給予時間反應
            except IndexError as ie:
                logging.error(f"CPLUS: 第 {i+1} 個按鈕處理出錯 (IndexError): {str(ie)} - 可能按鈕列表不足")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                driver.save_screenshot(f"house_index_error_{i+1}_{retry+1}_{timestamp}.png")
                with open(f"house_index_error_{i+1}_{retry+1}_{timestamp}.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                handle_popup(driver, wait)  # 失敗時再清視窗
                time.sleep(2)  # 等再試
            except Exception as e:
                logging.error(f"CPLUS: 第 {i+1} 個按鈕處理出錯: {str(e)}")
                handle_popup(driver, wait)  # 失敗時再清視窗
                time.sleep(1)
                if retry == 2:  # 最後一次記錄 debug
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    driver.save_screenshot(f"house_button_failure_{i+1}_{timestamp}.png")
                    with open(f"house_button_failure_{i+1}_{timestamp}.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
        if not success:
            logging.warning(f"CPLUS: 第 {i+1} 個 Excel 下載按鈕經過 3 次重試失敗")
            continue
    if new_files:
        logging.info(f"CPLUS: Housekeeping Reports 下載完成，共 {len(new_files)} 個文件，預期 {total_buttons} 個")
        if len(new_files) != total_buttons:
            logging.warning(f"CPLUS: 下載數 {len(new_files)} 不等於按鈕數 {total_buttons}，但繼續抽取現有檔案")  # 不 raise，繼續
    return new_files, len(new_files), total_buttons, report_files  # 無 new_files 也繼續
        
def process_cplus():
    driver = None
    downloaded_files = set()
    initial_files = set(os.listdir(cplus_download_dir))
    house_file_count = 0
    house_button_count = 0
    house_report_files = {} # 移出循環，累積跨重試
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
                        time.sleep(0.5) # 減到0.5s
                    if section_name != 'house':
                        new_files = section_func(driver, wait, initial_files)
                    else:
                        new_files, this_file_count, this_button_count, this_report_files = section_func(driver, wait, initial_files)
                        # 合併 report_files，選最新
                        for report_name, this_info in this_report_files.items():
                            if report_name in house_report_files:
                                if this_info['mod_time'] > house_report_files[report_name]['mod_time']:
                                    house_report_files[report_name] = this_info
                            else:
                                house_report_files[report_name] = this_info
                        house_file_count = this_file_count
                        house_button_count = this_button_count
                    downloaded_files.update(new_files)
                    initial_files.update(new_files)
                    success = True
                    break
                except Exception as e:
                    logging.error(f"CPLUS {section_name} 嘗試 {attempt+1}/{MAX_RETRIES} 失敗: {str(e)}")
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(0.5) # 減到0.5s
                        # 修改: 加刷新頁面或重新導航，避免內部崩潰殘留
                        try:
                            driver.refresh()
                        except:
                            pass
            if not success:
                logging.error(f"CPLUS {section_name} 經過 {MAX_RETRIES} 次嘗試失敗") # 不 raise，繼續抽取現有檔案
        return downloaded_files, house_file_count, house_button_count, driver, house_report_files
    except Exception as e:
        logging.error(f"CPLUS 總錯誤: {str(e)}")
        return downloaded_files, house_file_count, house_button_count, driver, house_report_files
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
                time.sleep(1) # 等待視窗出現
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
                success_logout = False
                for retry in range(2):  # 加重試 2 次
                    try:
                        # 點擊工具欄
                        logout_toolbar_barge = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
                        driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar_barge)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", logout_toolbar_barge)
                        logging.info("Barge: 工具欄點擊成功")
                        
                        # 點擊 Logout 選項
                        logout_span_xpath = "//div[contains(@class, 'mat-menu-panel')]//button//span[contains(text(), 'Logout')]"
                        logout_button_barge = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, logout_span_xpath)))
                        driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", logout_button_barge)
                        logging.info("Barge: Logout 選項點擊成功")
                        
                        # 等待頁面跳轉到登入頁（檢查 URL 或特定元素）
                        WebDriverWait(driver, 10).until(EC.url_contains("login"))
                        logging.info("Barge: 登出成功，頁面已跳轉到登入頁")
                        success_logout = True
                        break
                    except (TimeoutException, NoSuchElementException, WebDriverException) as te:
                        logging.warning(f"Barge: 登出嘗試 {retry+1}/2 失敗: {str(te)}，嘗試備用定位...")
                        try:
                            # 備用 Logout 定位
                            backup_logout_xpath = "//button[.//span[contains(text(), 'Logout')]]"
                            logout_button_barge = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, backup_logout_xpath)))
                            driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                            time.sleep(1)
                            driver.execute_script("arguments[0].click();", logout_button_barge)
                            logging.info("Barge: 備用 Logout 選項點擊成功")
                            
                            # 等待跳轉
                            WebDriverWait(driver, 10).until(EC.url_contains("login"))
                            logging.info("Barge: 登出成功，頁面已跳轉到登入頁")
                            success_logout = True
                            break
                        except Exception as be:
                            logging.warning(f"Barge: 備用登出失敗: {str(be)}")
                            if retry == 1:  # 最後一次記錄 debug
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                driver.save_screenshot(f"barge_logout_failure_{timestamp}.png")
                                with open(f"barge_logout_failure_{timestamp}.html", "w", encoding="utf-8") as f:
                                    f.write(driver.page_source)
                if not success_logout:
                    logging.error("Barge: 登出經過 2 次重試失敗，直接關閉 driver")
        except Exception as e:
            logging.error(f"Barge: 登出總失敗: {str(e)}")

def get_latest_file(download_dir, pattern):
    """
    取匹配pattern最新file：**優先冇'(1)'括號**，再最新mod_time。
    """
    try:
        all_files = [f for f in os.listdir(download_dir) 
                     if pattern in f and (f.endswith('.csv') or f.endswith('.xlsx'))]
        if not all_files:
            return None
        
        # **優先篩選：冇括號**
        no_bracket_files = [f for f in all_files if '( ' not in f and ' (' not in f]
        if no_bracket_files:
            # 無括號中選最新
            latest = max(no_bracket_files, key=lambda f: os.path.getmtime(os.path.join(download_dir, f)))
        else:
            # 全有括號，選最新
            latest = max(all_files, key=lambda f: os.path.getmtime(os.path.join(download_dir, f)))
        
        logging.info(f"✅ 選最新 [{pattern}]: {latest} (優先無括號)")
        return latest
    except Exception as e:
        logging.error(f"❌ get_latest_file ERR ({pattern}): {str(e)}")
        return None
        
def send_daily_email(house_report_files, house_button_count, cplus_dir, barge_dir):
    """
    全英Email：Subject大寫 + 日誌列所有附件file。
    """
    load_dotenv()
    try:
        smtp_server = os.environ.get('SMTP_SERVER', 'smtp.zoho.com')
        smtp_port = int(os.environ.get('SMTP_PORT', 587))
        sender_email = os.environ['ZOHO_EMAIL']
        sender_password = os.environ['ZOHO_PASSWORD']
        receiver_emails = os.environ.get('RECEIVER_EMAILS', 'paklun@ckline.com.hk').split(',')
        cc_emails = os.environ.get('CC_EMAILS', '').split(',') if os.environ.get('CC_EMAILS') else []
        dry_run = os.environ.get('DRY_RUN', 'False').lower() == 'true'
        gen_time = datetime.now().strftime('%d/%m/%Y %H:%M')

        # 最新file (優先無(1))
        movement_file = get_latest_file(cplus_dir, 'cntrMoveLog')
        onhand_file = get_latest_file(cplus_dir, 'data_')
        barge_file = get_latest_file(barge_dir, 'ContainerDetailReport')

        # House: 按mod_time排序(最新先)
        sorted_house = sorted(house_report_files.items(), key=lambda x: x[1]['mod_time'], reverse=True)
        house_download_count = len(sorted_house)

        # 附件清單
        attachments = []
        if movement_file: attachments.append((cplus_dir, movement_file))
        if onhand_file: attachments.append((cplus_dir, onhand_file))
        if barge_file: attachments.append((barge_dir, barge_file))
        for _, info in sorted_house:
            attachments.append((cplus_dir, info['file']))

        # **日誌：列所有附件file**
        attach_names = [f[1] for f in attachments]
        logging.info("📤 Email Attachments (%s files): %s", len(attach_names), ', '.join(attach_names))

        # HTML (全英)
        style = """
        <style>table{border-collapse:collapse;width:100%;font-family:Arial;font-size:14px;}
        th,td{border:1px solid #ddd;padding:10px;text-align:left;}
        th{background:#f2f2f2;font-weight:bold;}
        .sum{background:#e7f3ff;font-weight:bold;}
        </style>
        """
        num_house = len(sorted_house)
        body_html = f"""
        <html><head>{style}</head><body>
        <h2>HIT Daily Reports ({gen_time})</h2>
        <table>
        <thead><tr><th>Category</th><th>Report</th><th>File</th></tr></thead>
        <tbody>
        <tr><td rowspan="{2+num_house}">CPLUS</td><td>CONTAINER MOVEMENT</td><td>{movement_file}</td></tr>
        <tr><td>ONHAND CONTAINER LIST</td><td>{onhand_file}</td></tr>
        """
        for name, info in sorted_house:
            body_html += f'<tr><td>{name}</td><td>{info["file"]}</td></tr>'
        body_html += f"""
        <tr><td rowspan="1">BARGE</td><td>CONTAINER DETAIL REPORT</td><td>{barge_file}</td></tr>
        <tr class="sum"><td colspan="3">Housekeeping: {house_download_count}/{house_button_count} | Total Attachments: {len(attachments)}</td></tr>
        </tbody></table></body></html>
        """

        # Plain (全英)
        house_list = '\n'.join([f"  - {name}: {info['file']}" for name, info in sorted_house])
        plain_body = f"""HIT Daily Reports ({gen_time})

CPLUS:
- CONTAINER MOVEMENT: {movement_file}
- ONHAND CONTAINER LIST: {onhand_file}

Housekeeping Reports ({house_download_count}/{house_button_count}):
{house_list}

BARGE:
- CONTAINER DETAIL REPORT: {barge_file}

Total Attachments: {len(attachments)}
All files OK!
"""

        # Email (Subject **全大寫**)
        msg = MIMEMultipart('alternative')
        msg['From'] = sender_email
        msg['To'] = ', '.join(receiver_emails)
        if cc_emails: msg['Cc'] = ', '.join(cc_emails)
        msg['Subject'] = f"HIT DAILY REPORTS - {gen_time.upper()}"
        msg.attach(MIMEText(body_html, 'html'))
        msg.attach(MIMEText(plain_body, 'plain'))

        # 加附件
        for dir_path, file_name in attachments:
            file_path = os.path.join(dir_path, file_name)
            if os.path.exists(file_path):
                part = MIMEBase('application', 'octet-stream')
                with open(file_path, 'rb') as f:
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(f'Content-Disposition', f'attachment; filename="{file_name}"')
                msg.attach(part)

        if dry_run:
            logging.info("🧪 DRY RUN: Subject=%s | Files listed above", msg['Subject'])
            return

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_emails + cc_emails, msg.as_string())
        server.quit()
        logging.info("✅ Email Sent: %s files (listed above)", len(attachments))

    except Exception as e:
        logging.error("❌ Email ERR: %s", str(e))
        
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
    # **嚴格檢查：全齊才發**
    movement_file = get_latest_file(cplus_download_dir, 'cntrMoveLog')
    onhand_file = get_latest_file(cplus_download_dir, 'data_')
    barge_file = get_latest_file(barge_download_dir, 'ContainerDetailReport')
    
    movement_ok = movement_file is not None
    onhand_ok = onhand_file is not None
    barge_ok = barge_file is not None
    house_download_count = len(house_report_files)
    house_ok = (house_download_count == house_button_count)

    total_ok = int(movement_ok) + int(onhand_ok) + house_download_count + int(barge_ok)
    total_exp = 3 + house_button_count

    logging.info("📊 最終檢查: Movement=%s | OnHand=%s | Barge=%s | House=%s/%s | Total=%s/%s", 
                 '✓' if movement_ok else '✗', '✓' if onhand_ok else '✗', 
                 '✓' if barge_ok else '✗', house_download_count, house_button_count, total_ok, total_exp)

    # **總日誌：列** **所有** **下載file**（即使唔發）
    all_cplus_files = [f for f in os.listdir(cplus_download_dir) if f.endswith(('.csv', '.xlsx'))]
    all_barge_files = [f for f in os.listdir(barge_download_dir) if f.endswith(('.csv', '.xlsx'))]
    all_files = all_cplus_files + all_barge_files
    logging.info("📋 **所有** 下載 File (%s 個): %s", len(all_files), ', '.join(sorted(all_files)))

    if movement_ok and onhand_ok and barge_ok and house_ok:
        logging.info("🚀 全齊！發Email...")
        send_daily_email(house_report_files, house_button_count, cplus_download_dir, barge_download_dir)
    else:
        logging.warning("⚠️ 唔齊file，跳過Email！(需全✓)")

    logging.info("✅ 腳本完成")

if __name__ == "__main__":
    setup_environment()
    main()
