import os
import time
import shutil
import subprocess
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
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import logging
from dotenv import load_dotenv

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

cplus_download_dir = os.path.abspath("downloads_cplus")
barge_download_dir = os.path.abspath("downloads_barge")
MAX_RETRIES = 3
DOWNLOAD_TIMEOUT = 30  # 延長至 30 秒

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
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--disable-popup-blocking')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--no-first-run')
    chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
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

def wait_for_new_file(download_dir, initial_files, timeout=60):
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_files = set(f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx')))
        new_files = current_files - initial_files
        if new_files:
            return new_files
        if any(f.endswith('.crdownload') for f in os.listdir(download_dir)):
            logging.info("檢測到臨時文件，繼續等待...")
            time.sleep(5)
            continue
        time.sleep(1)
    logging.warning(f"未檢測到新文件，當前目錄內容: {os.listdir(download_dir)}")
    return set()

def handle_popup(driver, wait):
    try:
        error_div = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'System Error')]")))
        logging.info("檢測到 System Error popup")
        close_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Close')]")))
        ActionChains(driver).move_to_element(close_button).click().perform()
        logging.info("已點擊 Close 按鈕")
        WebDriverWait(driver, 5).until(EC.invisibility_of_element_located((By.XPATH, "//div[contains(text(), 'System Error')]")))
        logging.info("Popup 已消失")
        # 重試一次以確保處理
        handle_popup(driver, wait)
    except TimeoutException:
        logging.debug("無 popup 檢測到")

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

def process_cplus_movement(driver, wait, initial_files):
    logging.info("CPLUS: 直接前往 Container Movement Log...")
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
    time.sleep(2)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")))
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
            search_button_onhand = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search') or contains(@class, 'MuiButtonBase-root')]")))
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
    try:
        driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#root")))
        logging.info("CPLUS: Housekeeping Reports 頁面加載完成")
    except TimeoutException as e:
        logging.error(f"CPLUS: Housekeeping Reports 頁面加載超時: {str(e)}，當前 URL: {driver.current_url}")
        driver.save_screenshot("house_page_load_failure.png")
        with open("house_page_load_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: Housekeeping Reports 頁面加載失敗")

    logging.info("CPLUS: 等待表格加載...")
    try:
        rows = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table.MuiTable-root tbody tr")))
        if len(rows) == 0 or all(not row.text.strip() for row in rows):
            logging.debug("表格數據空或無效，刷新頁面...")
            driver.refresh()
            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table.MuiTable-root tbody tr")))
            rows = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table.MuiTable-root tbody tr")))
            if len(rows) < 6:
                logging.warning("刷新後表格數據仍不足，記錄頁面狀態...")
                driver.save_screenshot("house_load_failure.png")
                with open("house_load_failure.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                raise Exception("CPLUS: Housekeeping Reports 表格數據不足")
        logging.info(f"CPLUS: 表格加載完成，找到 {len(rows)} 行數據")
    except TimeoutException as e:
        logging.error(f"CPLUS: 表格加載失敗: {str(e)}")
        driver.save_screenshot("house_table_load_failure.png")
        with open("house_table_load_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: Housekeeping Reports 表格加載失敗")

    logging.info("CPLUS: 定位並點擊所有 Excel 下載按鈕...")
    local_initial = initial_files.copy()
    new_files = set()
    housekeep_prefixes = ['IE2_', 'DM1C_', 'IA17_', 'GA1_', 'IA5_', 'IA15_']
    
    # 使用更穩定的 CSS 選擇器定位按鈕
    try:
        excel_buttons = driver.find_elements(By.CSS_SELECTOR, "table.MuiTable-root tbody tr td button:not([disabled])")
        button_count = len(excel_buttons)
        logging.info(f"CPLUS: 找到 {button_count} 個 Excel 下載按鈕")
        
        # 備用 XPath 選擇器
        if button_count == 0:
            logging.debug("CPLUS: CSS 選擇器未找到按鈕，嘗試備用 XPath...")
            excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]")
            button_count = len(excel_buttons)
            logging.info(f"CPLUS: 備用 XPath 找到 {button_count} 個 Excel 下載按鈕")
    except Exception as e:
        logging.error(f"CPLUS: 定位 Excel 按鈕失敗: {str(e)}")
        driver.save_screenshot("house_buttons_load_failure.png")
        with open("house_buttons_load_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: 定位 Excel 下載按鈕失敗")

    for idx, button in enumerate(excel_buttons, 1):
        success = False
        report_name = "未知報告"
        for attempt in range(MAX_RETRIES):
            try:
                # 獲取報告名稱
                try:
                    report_name_elem = driver.find_element(By.CSS_SELECTOR, f"table.MuiTable-root tbody tr:nth-child({idx}) td:nth-child(3)")
                    report_name = report_name_elem.text.strip()
                    logging.info(f"CPLUS: 準備點擊第 {idx} 個 Excel 按鈕，報告名稱: {report_name}")
                except NoSuchElementException:
                    logging.debug(f"CPLUS: 無法獲取第 {idx} 個按鈕的報告名稱")

                # 滾動到按鈕並點擊
                driver.execute_script("arguments[0].scrollIntoView(true);", button)
                wait.until(EC.element_to_be_clickable(button))
                driver.execute_script("arguments[0].click();", button)
                logging.info(f"CPLUS: 第 {idx} 個 Excel 下載按鈕 (報告: {report_name}) JavaScript 點擊成功")

                # 處理可能的彈窗
                handle_popup(driver, wait)

                # 使用 ActionChains 再次點擊，確保成功
                ActionChains(driver).move_to_element(button).pause(0.5).click().perform()
                logging.info(f"CPLUS: 第 {idx} 個 Excel 下載按鈕 (報告: {report_name}) ActionChains 點擊成功")

                # 等待新文件
                temp_new = wait_for_new_file(cplus_download_dir, local_initial, timeout=60)
                if temp_new:
                    logging.info(f"CPLUS: 第 {idx} 個按鈕 (報告: {report_name}) 下載新文件: {', '.join(temp_new)}")
                    # 檢查文件名是否匹配預期前綴
                    matched_files = [f for f in temp_new if any(prefix in f for prefix in housekeep_prefixes)]
                    if matched_files:
                        logging.info(f"CPLUS: 報告 {report_name} 下載文件匹配預期: {matched_files}")
                    else:
                        logging.warning(f"CPLUS: 報告 {report_name} 下載文件不匹配預期前綴: {temp_new}")
                    local_initial.update(temp_new)
                    new_files.update(temp_new)
                    success = True
                    break
                else:
                    logging.warning(f"CPLUS: 第 {idx} 個按鈕 (報告: {report_name}) 未觸發新文件下載")
                    driver.save_screenshot(f"house_button_{idx}_failure.png")
                    with open(f"house_button_{idx}_failure.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
            except (TimeoutException, ElementClickInterceptedException) as e:
                logging.error(f"CPLUS: 第 {idx} 個 Excel 下載按鈕 (報告: {report_name}) 嘗試 {attempt+1}/{MAX_RETRIES} 失敗: {str(e)}")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(2)
            except Exception as e:
                logging.error(f"CPLUS: 第 {idx} 個 Excel 下載按鈕 (報告: {report_name}) 未知錯誤: {str(e)}")
                driver.save_screenshot(f"house_button_{idx}_failure.png")
                with open(f"house_button_{idx}_failure.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                break
        if not success:
            logging.warning(f"CPLUS: 第 {idx} 個 Excel 下載按鈕 (報告: {report_name}) 最終失敗")

    if new_files:
        logging.info(f"CPLUS: Housekeeping Reports 下載完成，共 {len(new_files)} 個文件，預期 {button_count} 個")
        # 檢查是否所有預期文件都下載
        downloaded_prefixes = {prefix for prefix in housekeep_prefixes if any(prefix in f for f in new_files)}
        missing_prefixes = set(housekeep_prefixes) - downloaded_prefixes
        if missing_prefixes:
            logging.warning(f"CPLUS: 缺少以下預期報告: {', '.join(missing_prefixes)}")
        else:
            logging.info("CPLUS: 所有預期報告文件均已下載")
        return new_files, len(new_files), button_count
    else:
        logging.error("CPLUS: Housekeeping Reports 未下載任何文件，記錄頁面狀態...")
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
    try:
        driver = webdriver.Chrome(options=get_chrome_options(cplus_download_dir))
        logging.info("CPLUS WebDriver 初始化成功")
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        wait = WebDriverWait(driver, 20)
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
                        time.sleep(5)
            if not success:
                logging.error(f"CPLUS {section_name} 經過 {MAX_RETRIES} 次嘗試失敗")
        return downloaded_files, house_file_count, house_button_count, driver
    except Exception as e:
        logging.error(f"CPLUS 總錯誤: {str(e)}")
        return downloaded_files, house_file_count, house_button_count, driver
    finally:
        try:
            if driver:
                logging.info("CPLUS: 嘗試登出...")
                logout_menu_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
                ActionChains(driver).move_to_element(logout_menu_button).click().perform()
                logging.info("CPLUS: 用戶菜單點擊成功")
                logout_option = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Logout')]")))
                ActionChains(driver).move_to_element(logout_option).click().perform()
                logging.info("CPLUS: Logout 選項點擊成功")
                time.sleep(2)
                # 加點擊 CLOSE
                try:
                    close_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="logout"]/div[3]/div/div[3]/button/span[1]')))
                    ActionChains(driver).move_to_element(close_button).click().perform()
                    logging.info("CPLUS: Logout 後 CLOSE 按鈕點擊成功")
                except TimeoutException:
                    logging.warning("CPLUS: Logout 後無 CLOSE 按鈕，跳過")
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
        wait = WebDriverWait(driver, 20)

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
                    logout_toolbar_barge = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
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
                    logout_button_barge = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, logout_span_xpath)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", logout_button_barge)
                    logging.info("Barge: Logout 選項點擊成功")
                except TimeoutException:
                    logging.debug("Barge: Logout 選項未找到，嘗試備用定位...")
                    backup_logout_xpath = "//button[.//span[contains(text(), 'Logout')]]"
                    logout_button_barge = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, backup_logout_xpath)))
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
    barge_driver = None

    # Process CPLUS sequentially
    cplus_result = process_cplus()
    cplus_files.update(cplus_result[0])
    house_file_count = cplus_result[1]
    house_button_count = cplus_result[2]
    cplus_driver = cplus_result[3]

    # Process Barge sequentially
    barge_result = process_barge()
    barge_files.update(barge_result[0])
    barge_driver = barge_result[1]

    logging.info("檢查所有下載文件...")
    downloaded_files = [f for f in os.listdir(cplus_download_dir) if f.endswith(('.csv', '.xlsx'))] + [f for f in os.listdir(barge_download_dir) if f.endswith(('.csv', '.xlsx'))]
    logging.info(f"總下載文件: {len(downloaded_files)} 個")
    for file in downloaded_files:
        logging.info(f"找到檔案: {file}")

    required_patterns = {'movement': 'cntrMoveLog', 'onhand': 'data_', 'barge': 'ContainerDetailReport'}
    housekeep_prefixes = ['IE2_', 'DM1C_', 'IA17_', 'GA1_', 'IA5_', 'IA15_']
    has_required = all(any(pattern in f for f in downloaded_files) for pattern in required_patterns.values())
    house_files = [f for f in downloaded_files if any(p in f for p in housekeep_prefixes)]
    house_download_count = len(house_files)
    house_ok = (house_button_count == 0) or (house_download_count >= house_button_count)

    if has_required and house_ok:
        logging.info("所有必須文件齊全，開始發送郵件...")
        try:
            smtp_server = os.environ.get('SMTP_SERVER', 'smtp.zoho.com')
            smtp_port = int(os.environ.get('SMTP_PORT', 587))
            sender_email = os.environ['ZOHO_EMAIL']
            sender_password = os.environ['ZOHO_PASSWORD']
            receiver_emails = os.environ.get('RECEIVER_EMAILS').split(',')
            cc_emails = os.environ.get('CC_EMAILS', '').split(',') if os.environ.get('CC_EMAILS') else []
            dry_run = os.environ.get('DRY_RUN', 'False').lower() == 'true'

            if dry_run:
                logging.info("Dry run 模式：只打印郵件內容，不發送。")

            house_report_names = [
                "REEFER CONTAINER MONITOR REPORT",
                "CONTAINER DAMAGE REPORT (LINE) ENTRY GATE + EXIT GATE",
                "CONTAINER LIST (ON HAND)",
                "CY - GATELOG",
                "CONTAINER LIST (DAMAGED)",
                "ACTIVE REEFER CONTAINER ON HAND LIST"
            ]
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
