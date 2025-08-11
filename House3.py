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
from functools import wraps

# 配置 logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("scraper.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# 全局變量
download_dir = os.path.abspath("downloads")
CPLUS_MOVEMENT_COUNT = 1  # Container Movement Log
CPLUS_ONHAND_COUNT = 1  # OnHandContainerList
BARGE_COUNT = 1  # Barge
MIN_HOUSE_BUTTONS = 6  # 最小預期 Housekeeping buttons，如果少於此，不 send email
MAX_RETRIES = 3

# 重試裝飾器
def retry_on_failure(max_retries=3, delay=5):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except TimeoutException as e:
                    logger.warning(f"{func.__name__} 時間超時 (嘗試 {attempt+1}/{max_retries}): {str(e)}")
                except ElementClickInterceptedException as e:
                    logger.warning(f"{func.__name__} 元素點擊被攔截 (嘗試 {attempt+1}/{max_retries}): {str(e)}")
                except NoSuchElementException as e:
                    logger.warning(f"{func.__name__} 元素未找到 (嘗試 {attempt+1}/{max_retries}): {str(e)}")
                except Exception as e:
                    logger.error(f"{func.__name__} 嘗試 {attempt+1}/{max_retries} 失敗: {str(e)}")
                    if attempt < max_retries - 1:
                        time.sleep(delay)
            raise Exception(f"{func.__name__} 經過 {max_retries} 次嘗試失敗")
        return wrapper
    return decorator

# 清空下載目錄
def clear_download_dir():
    if os.path.exists(download_dir):
        shutil.rmtree(download_dir)
    os.makedirs(download_dir)
    logger.info(f"創建下載目錄: {download_dir}")

# 確保環境準備
def setup_environment():
    try:
        result = subprocess.run(['which', 'chromium-browser'], capture_output=True, text=True)
        if result.returncode != 0:
            subprocess.run(['sudo', 'apt-get', 'update', '-qq'], check=True)
            subprocess.run(['sudo', 'apt-get', 'install', '-y', 'chromium-browser', 'chromium-chromedriver'], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            logger.info("Chromium 及 ChromeDriver 已安裝")
        else:
            logger.info("Chromium 及 ChromeDriver 已存在，跳過安裝")
        result = subprocess.run(['pip', 'show', 'selenium'], capture_output=True, text=True)
        if "selenium" not in result.stdout or "webdriver-manager" not in subprocess.run(['pip', 'show', 'webdriver-manager'], capture_output=True, text=True).stdout:
            subprocess.run(['pip', 'install', 'selenium', 'webdriver-manager'], check=True)
            logger.info("Selenium 及 WebDriver Manager 已安裝")
        else:
            logger.info("Selenium 及 WebDriver Manager 已存在，跳過安裝")
    except subprocess.CalledProcessError as e:
        logger.error(f"環境準備失敗: {e}")
        raise

# 設置 Chrome 選項
def get_chrome_options():
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

# 檢查新文件出現，並過濾 size > 0
def wait_for_new_file(initial_files, start_time, timeout=30, expected_pattern=None):  # 增加 timeout 到 30s
    end_time = time.time() + timeout
    while time.time() < end_time:
        current_files = set(f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx')))
        new_files = current_files - initial_files
        if new_files:
            filtered_files = set()
            for file in new_files:
                file_path = os.path.join(download_dir, file)
                if os.path.getmtime(file_path) >= start_time and os.path.getsize(file_path) > 0:  # 加 size check
                    if expected_pattern is None or expected_pattern in file:
                        filtered_files.add(file)
            if filtered_files:
                return filtered_files
        time.sleep(0.5)
    return set()

# CPLUS 登入
@retry_on_failure(max_retries=MAX_RETRIES)
def cplus_login(driver, wait):
    logger.info("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/")
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    logger.info(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}")
    time.sleep(2)
    logger.info("CPLUS: 點擊登錄前按鈕...")
    login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
    ActionChains(driver).move_to_element(login_button_pre).click().perform()
    logger.info("CPLUS: 登錄前按鈕點擊成功")
    time.sleep(2)
    logger.info("CPLUS: 輸入 COMPANY CODE...")
    company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    company_code_field.send_keys("CKL")
    logger.info("CPLUS: COMPANY CODE 輸入完成")
    time.sleep(1)
    logger.info("CPLUS: 輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
    user_id_field.send_keys("KEN")
    logger.info("CPLUS: USER ID 輸入完成")
    time.sleep(1)
    logger.info("CPLUS: 輸入 PASSWORD...")
    password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
    password_field.send_keys(os.environ.get('SITE_PASSWORD'))  # 移除 fallback，強制用 env
    logger.info("CPLUS: PASSWORD 輸入完成")
    time.sleep(1)
    logger.info("CPLUS: 點擊 LOGIN 按鈕...")
    login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
    ActionChains(driver).move_to_element(login_button).click().perform()
    logger.info("CPLUS: LOGIN 按鈕點擊成功")
    time.sleep(2)

# CPLUS Container Movement Log
@retry_on_failure(max_retries=MAX_RETRIES)
def process_cplus_movement(driver, wait, initial_files):
    logger.info("CPLUS: 直接前往 Container Movement Log...")
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
    time.sleep(2)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")))
    logger.info("CPLUS: Container Movement Log 頁面加載完成")
    logger.info("CPLUS: 點擊 Search...")
    local_initial = initial_files.copy()
    start_time = time.time()
    for attempt in range(2):
        try:
            search_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")))
            ActionChains(driver).move_to_element(search_button).click().perform()
            logger.info("CPLUS: Search 按鈕點擊成功")
            break
        except TimeoutException:
            logger.warning(f"CPLUS: Search 按鈕未找到，嘗試備用定位 {attempt+1}/2...")
            try:
                search_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and .//span[contains(text(), 'Search')]]")))
                ActionChains(driver).move_to_element(search_button).click().perform()
                logger.info("CPLUS: 備用 Search 按鈕 1 點擊成功")
                break
            except TimeoutException:
                logger.warning(f"CPLUS: 備用 Search 按鈕 1 失敗，嘗試備用定位 2 (嘗試 {attempt+1}/2)...")
                try:
                    search_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]")))
                    ActionChains(driver).move_to_element(search_button).click().perform()
                    logger.info("CPLUS: 備用 Search 按鈕 2 點擊成功")
                    break
                except TimeoutException:
                    logger.warning(f"CPLUS: 備用 Search 按鈕 2 失敗 (嘗試 {attempt+1}/2)")
    else:
        raise Exception("CPLUS: Container Movement Log Search 按鈕點擊失敗")
    logger.info("CPLUS: 點擊 Download...")
    for attempt in range(2):
        try:
            download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button")))
            ActionChains(driver).move_to_element(download_button).click().perform()
            logger.info("CPLUS: Download 按鈕點擊成功")
            time.sleep(0.5)
            try:
                driver.execute_script("arguments[0].click();", download_button)
                logger.info("CPLUS: Download 按鈕 JavaScript 點擊成功")
            except Exception as js_e:
                logger.warning(f"CPLUS: Download 按鈕 JavaScript 點擊失敗: {str(js_e)}")
            time.sleep(0.5)
            break
        except Exception as e:
            logger.warning(f"CPLUS: Download 按鈕點擊失敗 (嘗試 {attempt+1}/2): {str(e)}")
            time.sleep(0.5)
    else:
        raise Exception("CPLUS: Container Movement Log Download 按鈕點擊失敗")
    new_files = wait_for_new_file(local_initial, start_time, timeout=30, expected_pattern=None)  # 增加 timeout
    filtered_files = {f for f in new_files if "ContainerDetailReport" not in f}
    if len(filtered_files) >= CPLUS_MOVEMENT_COUNT:
        logger.info(f"CPLUS: Container Movement Log 下載完成，檔案位於: {download_dir}")
        for file in filtered_files:
            logger.info(f"CPLUS: 新下載檔案: {file}")
        return filtered_files
    else:
        raise Exception("CPLUS: Container Movement Log 未觸發足夠新文件下載")

# CPLUS OnHandContainerList
@retry_on_failure(max_retries=MAX_RETRIES)
def process_cplus_onhand(driver, wait, initial_files):
    logger.info("CPLUS: 前往 OnHandContainerList 頁面...")
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/OnHandContainerList")
    time.sleep(2)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    logger.info("CPLUS: OnHandContainerList 頁面加載完成")
    logger.info("CPLUS: 點擊 Search...")
    local_initial = initial_files.copy()
    start_time = time.time()
    try:
        search_button_onhand = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")))  # 增加 wait 到 20s
        ActionChains(driver).move_to_element(search_button_onhand).click().perform()
        logger.info("CPLUS: Search 按鈕點擊成功")
    except TimeoutException:
        logger.warning("CPLUS: Search 按鈕未找到，嘗試備用定位...")
        search_button_onhand = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]")))
        ActionChains(driver).move_to_element(search_button_onhand).click().perform()
        logger.info("CPLUS: 備用 Search 按鈕點擊成功")
    time.sleep(0.5)
    logger.info("CPLUS: 點擊 Export...")
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button")))
    ActionChains(driver).move_to_element(export_button).click().perform()
    logger.info("CPLUS: Export 按鈕點擊成功")
    time.sleep(0.5)
    logger.info("CPLUS: 點擊 Export as CSV...")
    export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']")))
    ActionChains(driver).move_to_element(export_csv_button).click().perform()
    logger.info("CPLUS: Export as CSV 按鈕點擊成功")
    time.sleep(0.5)
    new_files = wait_for_new_file(local_initial, start_time, timeout=30, expected_pattern=None)
    filtered_files = {f for f in new_files if "ContainerDetailReport" not in f}
    if len(filtered_files) >= CPLUS_ONHAND_COUNT:
        logger.info(f"CPLUS: OnHandContainerList 下載完成，檔案位於: {download_dir}")
        for file in filtered_files:
            logger.info(f"CPLUS: 新下載檔案: {file}")
        return filtered_files
    else:
        raise Exception("CPLUS: OnHandContainerList 未觸發足夠新文件下載")

# CPLUS Housekeeping Reports
@retry_on_failure(max_retries=MAX_RETRIES)
def process_cplus_house(driver, wait, initial_files):
    logger.info("CPLUS: 前往 Housekeeping Reports 頁面...")
    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
    time.sleep(2)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    logger.info("CPLUS: Housekeeping Reports 頁面加載完成")
    logger.info("CPLUS: 等待表格加載...")
    for attempt in range(3):
        try:
            wait = WebDriverWait(driver, 20)  # 增加到 20s
            rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
            logger.info(f"CPLUS: 表格加載完成，找到 {len(rows)} 行")
            break
        except TimeoutException:
            logger.warning(f"CPLUS: 表格未加載，嘗試刷新頁面 (嘗試 {attempt+1}/3)...")
            driver.refresh()
            time.sleep(2)
    else:
        raise Exception("CPLUS: Housekeeping Reports 表格加載失敗")
    logger.info("CPLUS: 定位並點擊所有 Excel 下載按鈕...")
    local_initial = initial_files.copy()
    start_time = time.time()
    xpath_attempts = [
        "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]//svg//path[contains(@d, 'M20 12l')]",  # 新加: 用 download icon path d
        "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]//svg[@viewBox='0 0 24 24']//path[@fill='#036e11']",
        "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)]",
        "//table//tbody//tr//td//button[.//span[contains(text(), 'Download')]]",
        "//button[contains(@class, 'MuiButtonBase-root') and not(@disabled)]//svg[@data-testid='DownloadIcon']",
        "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button",
        "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td//button[contains(@aria-label, 'download') or contains(@title, 'download')]"  # 新加: aria-label or title
    ]
    excel_buttons = []
    for xpath in xpath_attempts:
        try:
            excel_buttons = driver.find_elements(By.XPATH, xpath)
            if excel_buttons:
                logger.info(f"CPLUS: 使用 XPath '{xpath}' 找到 {len(excel_buttons)} 個 Excel 下載按鈕")
                break
        except Exception as e:
            logger.warning(f"CPLUS: XPath '{xpath}' 定位失敗: {str(e)}")
    button_count = len(excel_buttons)
    logger.info(f"CPLUS: 最終找到 {button_count} 個 Excel 下載按鈕")
    if button_count == 0:
        logger.warning("CPLUS: 未找到任何 Excel 下載按鈕，記錄 page source 以 debug")
        with open("housekeep_page_source.html", "w") as f:
            f.write(driver.page_source)
        logger.info("CPLUS: page source 保存到 housekeep_page_source.html")
        return set(), 0, 0
    new_files = set()
    for idx in range(button_count):
        success = False
        for button_attempt in range(3):
            try:
                button_xpath = f"(//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)])[{idx+1}]"
                button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
                driver.execute_script("arguments[0].scrollIntoView(true);", button)
                time.sleep(0.5)
                try:
                    report_name = driver.find_element(By.XPATH, f"//table[contains(@class, 'MuiTable-root')]//tbody//tr[{idx+1}]//td[3]").text
                    logger.info(f"CPLUS: 準備點擊第 {idx+1} 個 Excel 按鈕 (嘗試 {button_attempt+1}/3)，報告名稱: {report_name}")
                except:
                    logger.warning(f"CPLUS: 無法獲取第 {idx+1} 個按鈕的報告名稱 (嘗試 {button_attempt+1}/3)")
                ActionChains(driver).move_to_element(button).pause(0.5).click().perform()
                logger.info(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕點擊成功 (嘗試 {button_attempt+1}/3)")
                time.sleep(2)  # 加等 2s 確保下載
                temp_new = wait_for_new_file(local_initial, start_time, timeout=30)
                filtered_files = {f for f in temp_new if "ContainerDetailReport" not in f}
                if filtered_files:
                    logger.info(f"CPLUS: 第 {idx+1} 個按鈕下載新文件: {', '.join(filtered_files)} (嘗試 {button_attempt+1}/3)")
                    local_initial.update(filtered_files)
                    new_files.update(filtered_files)
                    success = True
                    break
                else:
                    logger.warning(f"CPLUS: 第 {idx+1} 個按鈕未觸發新文件下載 (嘗試 {button_attempt+1}/3)")
                driver.execute_script("arguments[0].click();", button)
                logger.info(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕 JavaScript 點擊成功 (嘗試 {button_attempt+1}/3)")
                time.sleep(2)
                temp_new = wait_for_new_file(local_initial, start_time, timeout=30)
                filtered_files = {f for f in temp_new if "ContainerDetailReport" not in f}
                if filtered_files:
                    logger.info(f"CPLUS: 第 {idx+1} 個按鈕下載新文件 (JavaScript): {', '.join(filtered_files)} (嘗試 {button_attempt+1}/3)")
                    local_initial.update(filtered_files)
                    new_files.update(filtered_files)
                    success = True
                    break
                else:
                    logger.warning(f"CPLUS: 第 {idx+1} 個按鈕 JavaScript 未觸發新文件下載 (嘗試 {button_attempt+1}/3)")
            except Exception as e:
                logger.warning(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕點擊失敗 (嘗試 {button_attempt+1}/3): {str(e)}")
        if not success:
            logger.warning(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕經過 3 次嘗試失敗")
    if len(new_files) != button_count:
        logger.warning(f"CPLUS: Housekeeping Reports 下載文件數量 ({len(new_files)}) 不等於按鈕數量 ({button_count})")
        raise Exception(f"CPLUS: Housekeeping Reports 下載文件數量不正確，預期 {button_count} 個")
    logger.info(f"CPLUS: Housekeeping Reports 下載完成，共 {len(new_files)} 個文件，預期 {button_count} 個")
    return new_files, len(new_files), button_count

# CPLUS 操作
def process_cplus():
    driver = None
    downloaded_files = set()
    initial_files = set()
    house_file_count = 0
    house_button_count = 0
    try:
        driver = webdriver.Chrome(options=get_chrome_options())
        logger.info("CPLUS WebDriver 初始化成功")
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
                    logger.info(f"CPLUS: 開始處理 {section_name}，嘗試 {attempt+1}/{MAX_RETRIES}")
                    if section_name == 'house':
                        new_files, count, button_count = section_func(driver, wait, initial_files)
                        house_file_count = count
                        house_button_count = button_count
                    else:
                        new_files = section_func(driver, wait, initial_files)
                    downloaded_files.update(new_files)
                    initial_files.update(new_files)
                    success = True
                    logger.info(f"CPLUS: {section_name} 處理成功，新增文件: {new_files}")
                    break
                except Exception as e:
                    logger.error(f"CPLUS: {section_name} 嘗試 {attempt+1}/{MAX_RETRIES} 失敗: {str(e)}")
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(5)
            if not success:
                logger.error(f"CPLUS: {section_name} 經過 {MAX_RETRIES} 次嘗試失敗")
                return set(), 0, 0  # 如果 any fail，返回 empty
        return downloaded_files, house_file_count, house_button_count
    except Exception as e:
        logger.error(f"CPLUS 總錯誤: {str(e)}")
        return set(), 0, 0
    finally:
        if driver:
            try:
                logger.info("CPLUS: 嘗試登出...")
                logout_menu_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
                ActionChains(driver).move_to_element(logout_menu_button).click().perform()
                logger.info("CPLUS: 用戶菜單點擊成功")
                logout_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Logout')]")))
                ActionChains(driver).move_to_element(logout_option).click().perform()
                logger.info("CPLUS: Logout 選項點擊成功")
                time.sleep(2)
                logger.info("CPLUS: 點擊 Close 按鈕完成登出...")
                try:
                    close_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.MuiButtonBase-root span.MuiButton-label")))
                    ActionChains(driver).move_to_element(close_button).click().perform()
                    logger.info("CPLUS: Close 按鈕點擊成功 (CSS)")
                except TimeoutException:
                    logger.warning("CPLUS: CSS 選擇器未找到 Close 按鈕，嘗試 XPath...")
                    try:
                        close_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and .//span[contains(text(), 'Close')]]")))
                        ActionChains(driver).move_to_element(close_button).click().perform()
                        logger.info("CPLUS: Close 按鈕點擊成功 (XPath)")
                    except TimeoutException:
                        logger.warning("CPLUS: XPath 選擇器未找到 Close 按鈕，嘗試備用定位...")
                        try:
                            close_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Close')]")))
                            ActionChains(driver).move_to_element(close_button).click().perform()
                            logger.info("CPLUS: Close 按鈕點擊成功 (備用 XPath)")
                        except TimeoutException as e:
                            logger.warning(f"CPLUS: Close 按鈕點擊失敗: {str(e)}")
                time.sleep(2)
                try:
                    driver.get("https://cplus.hit.com.hk/frontpage/#/")
                    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
                    logger.info("CPLUS: 登出成功，回到登入頁")
                except TimeoutException:
                    logger.warning("CPLUS: 登出驗證失敗，可能未完全登出")
            except Exception as e:
                logger.error(f"CPLUS: 登出失敗: {str(e)}")
            finally:
                driver.quit()
                logger.info("CPLUS WebDriver 關閉")

# Barge 登入
@retry_on_failure(max_retries=MAX_RETRIES)
def barge_login(driver, wait):
    logger.info("Barge: 嘗試打開網站 https://barge.oneport.com/login...")
    driver.get("https://barge.oneport.com/login")
    logger.info(f"Barge: 網站已成功打開，當前 URL: {driver.current_url}")
    time.sleep(3)
    logger.info("Barge: 輸入 COMPANY ID...")
    company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Company ID' or contains(@id, 'mat-input-0')]")))
    company_id_field.send_keys("CKL")
    logger.info("Barge: COMPANY ID 輸入完成")
    time.sleep(1)
    logger.info("Barge: 輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='User ID' or contains(@id, 'mat-input-1')]")
    user_id_field.send_keys("barge")
    logger.info("Barge: USER ID 輸入完成")
    time.sleep(1)
    logger.info("Barge: 輸入 PW...")
    password_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Password' or contains(@id, 'mat-input-2')]")
    password_field.send_keys(os.environ.get('BARGE_PASSWORD', '123456'))  # 如果有 env，用 env，否則 fallback (但建議 set env)
    logger.info("Barge: PW 輸入完成")
    time.sleep(1)
    logger.info("Barge: 點擊 LOGIN 按鈕...")
    login_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'LOGIN') or contains(@class, 'mat-raised-button')]")))
    ActionChains(driver).move_to_element(login_button_barge).click().perform()
    logger.info("Barge: LOGIN 按鈕點擊成功")
    time.sleep(3)

# Barge 下載部分
@retry_on_failure(max_retries=MAX_RETRIES)
def process_barge_download(driver, wait, initial_files):
    logger.info("Barge: 直接前往 https://barge.oneport.com/downloadReport...")
    driver.get("https://barge.oneport.com/downloadReport")
    time.sleep(3)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    logger.info("Barge: downloadReport 頁面加載完成")
    logger.info("Barge: 選擇 Report Type...")
    start_time = time.time()
    report_type_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-form-field[.//mat-label[contains(text(), 'Report Type')]]//div[contains(@class, 'mat-select-trigger')]")))
    ActionChains(driver).move_to_element(report_type_trigger).click().perform()
    logger.info("Barge: Report Type 選擇開始")
    time.sleep(2)
    logger.info("Barge: 點擊 Container Detail...")
    container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-option//span[contains(text(), 'Container Detail')]")))
    ActionChains(driver).move_to_element(container_detail_option).click().perform()
    logger.info("Barge: Container Detail 點擊成功")
    time.sleep(2)
    logger.info("Barge: 點擊 Download...")
    local_initial = initial_files.copy()
    download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Download']]")))
    ActionChains(driver).move_to_element(download_button_barge).click().perform()
    logger.info("Barge: Download 按鈕點擊成功")
    new_files = wait_for_new_file(local_initial, start_time, timeout=30, expected_pattern="ContainerDetailReport")
    filtered_files = {f for f in new_files if "ContainerDetailReport" in f}
    if len(filtered_files) >= BARGE_COUNT:
        logger.info(f"Barge: Container Detail 下載完成，檔案位於: {download_dir}")
        for file in filtered_files:
            logger.info(f"Barge: 新下載檔案: {file}")
        return filtered_files
    else:
        raise Exception("Barge: Container Detail 未觸發足夠新文件下載")

# Barge 操作
def process_barge():
    driver = None
    downloaded_files = set()
    try:
        driver = webdriver.Chrome(options=get_chrome_options())
        logger.info("Barge WebDriver 初始化成功")
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        wait = WebDriverWait(driver, 10)
        barge_login(driver, wait)
        success = False
        for attempt in range(MAX_RETRIES):
            try:
                new_files = process_barge_download(driver, wait, downloaded_files)
                downloaded_files.update(new_files)
                success = True
                break
            except Exception as e:
                logger.error(f"Barge 下載嘗試 {attempt+1}/{MAX_RETRIES} 失敗: {str(e)}")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(5)
        if not success:
            logger.error(f"Barge 下載經過 {MAX_RETRIES} 次嘗試失敗")
            return set()  # fail 返回 empty
        return downloaded_files
    except Exception as e:
        logger.error(f"Barge 總錯誤: {str(e)}")
        return set()
    finally:
        try:
            if driver:
                logger.info("Barge: 點擊工具欄進行登出...")
                try:
                    logout_toolbar_barge = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar_barge)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", logout_toolbar_barge)
                    logger.info("Barge: 工具欄點擊成功")
                except TimeoutException:
                    logger.warning("Barge: 主工具欄登出按鈕未找到，嘗試備用定位...")
                    raise
                time.sleep(2)
                logger.info("Barge: 點擊 Logout 選項...")
                try:
                    logout_span_xpath = "//div[contains(@class, 'mat-menu-panel')]//button//span[contains(text(), 'Logout')]"
                    logout_button_barge = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, logout_span_xpath)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", logout_button_barge)
                    logger.info("Barge: Logout 選項點擊成功")
                except TimeoutException:
                    logger.warning("Barge: Logout 選項未找到，嘗試備用定位...")
                    backup_logout_xpath = "//button[.//span[contains(text(), 'Logout')]]"
                    logout_button_barge = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, backup_logout_xpath)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", logout_button_barge)
                    logger.info("Barge: 備用 Logout 選項點擊成功")
                time.sleep(5)
        except Exception as e:
            logger.error(f"Barge: 登出失敗: {str(e)}")
        if driver:
            driver.quit()
            logger.info("Barge WebDriver 關閉")

# 主函數
def main():
    clear_download_dir()
    cplus_files, house_file_count, house_button_count = process_cplus()  # 先 CPLUS
    barge_files = process_barge()  # 再 Barge，sequential
    logger.info("檢查所有下載文件...")
    downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
    expected_file_count = CPLUS_MOVEMENT_COUNT + CPLUS_ONHAND_COUNT + house_button_count + BARGE_COUNT
    logger.info(f"預期文件數量: {expected_file_count} (Movement: {CPLUS_MOVEMENT_COUNT}, OnHand: {CPLUS_ONHAND_COUNT}, Housekeeping: {house_button_count}, Barge: {BARGE_COUNT})")
    # 檢查 Housekeeping 文件數量是否匹配按鈕數量，且 >= MIN_HOUSE_BUTTONS
    if house_file_count != house_button_count or house_button_count < MIN_HOUSE_BUTTONS:
        logger.error(f"Housekeeping Reports 下載文件數量（{house_file_count}）不等於按鈕數量（{house_button_count}），或少於最小預期 {MIN_HOUSE_BUTTONS}，放棄發送郵件")
        logger.info("下載文件列表：" + str(downloaded_files))
        return
    # 檢查所有文件數量
    logger.info(f"總共下載 {len(downloaded_files)} 個文件:")
    for file in downloaded_files:
        logger.info(f"找到檔案: {file}")
    # 檢查 CPLUS 和 Barge 文件數量
    cplus_file_count = len([f for f in downloaded_files if "ContainerDetailReport" not in f])
    barge_file_count = len([f for f in downloaded_files if "ContainerDetailReport" in f])
    if cplus_file_count < (CPLUS_MOVEMENT_COUNT + CPLUS_ONHAND_COUNT + house_button_count) or barge_file_count < BARGE_COUNT:
        logger.error(f"CPLUS 文件數量（{cplus_file_count}）或 Barge 文件數量（{barge_file_count}）不足，預期 CPLUS: {CPLUS_MOVEMENT_COUNT + CPLUS_ONHAND_COUNT + house_button_count}，Barge: {BARGE_COUNT}")
        logger.info("下載文件列表：" + str(downloaded_files))
        return
    if len(downloaded_files) == expected_file_count:
        logger.info(f"所有下載完成，檔案位於: {download_dir}")
        logger.info("開始發送郵件...")
        try:
            smtp_server = 'smtp.zoho.com'
            smtp_port = 587
            sender_email = os.environ.get('ZOHO_EMAIL')
            sender_password = os.environ.get('ZOHO_PASSWORD')
            if not sender_email or not sender_password:
                raise ValueError("ZOHO_EMAIL 或 ZOHO_PASSWORD env var 未 set")
            receiver_email = 'paklun@ckline.com.hk'
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = receiver_email
            msg['Subject'] = f"[TESTING] HIT DAILY {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            body = "Attached are the daily reports downloaded from CPLUS (Container Movement Log, OnHand Container List, and Housekeeping Reports) and Barge (Container Detail)."
            msg.attach(MIMEText(body, 'plain'))
            for file in downloaded_files:
                file_path = os.path.join(download_dir, file)
                attachment = MIMEBase('application', 'octet-stream')
                attachment.set_payload(open(file_path, 'rb').read())
                encoders.encode_base64(attachment)
                attachment.add_header('Content-Disposition', f'attachment; filename={file}')
                msg.attach(attachment)
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
            server.quit()
            logger.info("郵件發送成功!")
        except Exception as e:
            logger.error(f"郵件發送失敗: {str(e)}")
    else:
        logger.error(f"總下載文件數量不足（{len(downloaded_files)}/{expected_file_count}），放棄發送郵件")
        logger.info("下載文件列表：" + str(downloaded_files))
    logger.info("腳本完成")

if __name__ == "__main__":
    setup_environment()
    main()
