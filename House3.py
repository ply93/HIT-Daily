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
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import logging
from functools import wraps
from concurrent.futures import ThreadPoolExecutor
import tenacity

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
MAX_RETRIES = 3

# 重試裝飾器
def retry_on_failure(max_retries=3, delay=5):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
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

# 檢查新文件出現（優化：動態超時）
def wait_for_new_file(initial_files, start_time, timeout=10, expected_pattern=None):
    max_timeout = max(10, timeout * 1.5)  # 動態增加 50% 緩衝
    end_time = time.time() + max_timeout
    while time.time() < end_time:
        current_files = set(f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx')))
        new_files = current_files - initial_files
        if new_files:
            filtered_files = set()
            for file in new_files:
                file_path = os.path.join(download_dir, file)
                if os.path.getmtime(file_path) >= start_time:
                    if expected_pattern is None or expected_pattern in file:
                        filtered_files.add(file)
            if filtered_files:
                return filtered_files
        time.sleep(0.5)
    logger.warning(f"未在 {max_timeout} 秒內檢測到新文件")
    return set()

# CPLUS 登入
@retry_on_failure(max_retries=MAX_RETRIES)
def cplus_login(driver, wait):
    logger.info("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/")
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    logger.info(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
    login_button_pre = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")
    ActionChains(driver).move_to_element(login_button_pre).click().perform()
    logger.info("CPLUS: 登錄前按鈕點擊成功")
    company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    company_code_field.send_keys("CKL")
    logger.info("CPLUS: COMPANY CODE 輸入完成")
    user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
    user_id_field.send_keys("KEN")
    logger.info("CPLUS: USER ID 輸入完成")
    password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
    password_field.send_keys(os.environ.get('SITE_PASSWORD'))
    logger.info("CPLUS: PASSWORD 輸入完成")
    login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
    ActionChains(driver).move_to_element(login_button).click().perform()
    logger.info("CPLUS: LOGIN 按鈕點擊成功")

# CPLUS Container Movement Log
@retry_on_failure(max_retries=MAX_RETRIES)
def process_cplus_movement(driver, wait, initial_files):
    logger.info("CPLUS: 直接前往 Container Movement Log...")
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")))
    logger.info("CPLUS: Container Movement Log 頁面加載完成")
    local_initial = initial_files.copy()
    start_time = time.time()
    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Search']]")))
    ActionChains(driver).move_to_element(search_button).click().perform()
    logger.info("CPLUS: Search 按鈕點擊成功")
    download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button")))
    ActionChains(driver).move_to_element(download_button).click().perform()
    logger.info("CPLUS: Download 按鈕點擊成功")
    new_files = wait_for_new_file(local_initial, start_time, timeout=10)
    filtered_files = {f for f in new_files if "ContainerDetailReport" not in f}
    if len(filtered_files) >= CPLUS_MOVEMENT_COUNT:
        logger.info(f"CPLUS: Container Movement Log 下載完成")
        return filtered_files
    else:
        raise Exception("CPLUS: Container Movement Log 未觸發足夠新文件下載")

# CPLUS OnHandContainerList
@retry_on_failure(max_retries=MAX_RETRIES)
def process_cplus_onhand(driver, wait, initial_files):
    logger.info("CPLUS: 前往 OnHandContainerList 頁面...")
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/OnHandContainerList")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    logger.info("CPLUS: OnHandContainerList 頁面加載完成")
    local_initial = initial_files.copy()
    start_time = time.time()
    search_button_onhand = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Search']]")))
    ActionChains(driver).move_to_element(search_button_onhand).click().perform()
    logger.info("CPLUS: Search 按鈕點擊成功")
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button")))
    ActionChains(driver).move_to_element(export_button).click().perform()
    logger.info("CPLUS: Export 按鈕點擊成功")
    export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']")))
    ActionChains(driver).move_to_element(export_csv_button).click().perform()
    logger.info("CPLUS: Export as CSV 按鈕點擊成功")
    new_files = wait_for_new_file(local_initial, start_time, timeout=10)
    filtered_files = {f for f in new_files if "ContainerDetailReport" not in f}
    if len(filtered_files) >= CPLUS_ONHAND_COUNT:
        logger.info(f"CPLUS: OnHandContainerList 下載完成")
        return filtered_files
    else:
        raise Exception("CPLUS: OnHandContainerList 未觸發足夠新文件下載")

# CPLUS Housekeeping Reports
@retry_on_failure(max_retries=MAX_RETRIES)
def process_cplus_house(driver, wait, initial_files):
    logger.info("CPLUS: 前往 Housekeeping Reports 頁面...")
    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
    logger.info("CPLUS: Housekeeping Reports 頁面加載完成")
    local_initial = initial_files.copy()
    start_time = time.time()
    excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]")
    button_count = len(excel_buttons)
    logger.info(f"CPLUS: 找到 {button_count} 個 Excel 下載按鈕")
    new_files = set()
    for idx, button in enumerate(excel_buttons):
        driver.execute_script("arguments[0].scrollIntoView(true);", button)
        ActionChains(driver).move_to_element(button).click().perform()
        temp_new = wait_for_new_file(local_initial, start_time, timeout=10)
        filtered_files = {f for f in temp_new if "ContainerDetailReport" not in f}
        if filtered_files:
            new_files.update(filtered_files)
            local_initial.update(filtered_files)
    if len(new_files) == button_count:
        logger.info(f"CPLUS: Housekeeping Reports 下載完成，共 {len(new_files)} 個文件")
        return new_files, len(new_files), button_count
    else:
        raise Exception(f"CPLUS: Housekeeping Reports 下載文件數量不正確，預期 {button_count} 個")

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
            process_cplus_movement,
            process_cplus_onhand,
            process_cplus_house
        ]
        for section_func in sections:
            if section_func == process_cplus_house:
                new_files, count, button_count = section_func(driver, wait, initial_files)
                house_file_count = count
                house_button_count = button_count
            else:
                new_files = section_func(driver, wait, initial_files)
            downloaded_files.update(new_files)
            initial_files.update(new_files)
        return downloaded_files, house_file_count, house_button_count
    finally:
        if driver:
            try:
                logout_menu_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
                ActionChains(driver).move_to_element(logout_menu_button).click().perform()
                logout_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Logout')]")))
                ActionChains(driver).move_to_element(logout_option).click().perform()
                close_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Close')]")))
                ActionChains(driver).move_to_element(close_button).click().perform()
                logger.info("CPLUS: 登出成功")
            except Exception as e:
                logger.error(f"CPLUS: 登出失敗: {str(e)}")
            driver.quit()
            logger.info("CPLUS WebDriver 關閉")

# Barge 登入
@retry_on_failure(max_retries=MAX_RETRIES)
def barge_login(driver, wait):
    logger.info("Barge: 嘗試打開網站 https://barge.oneport.com/login...")
    driver.get("https://barge.oneport.com/login")
    logger.info(f"Barge: 網站已成功打開，當前 URL: {driver.current_url}")
    company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Company ID' or contains(@id, 'mat-input-0')]")))
    company_id_field.send_keys("CKL")
    logger.info("Barge: COMPANY ID 輸入完成")
    user_id_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='User ID' or contains(@id, 'mat-input-1')]")
    user_id_field.send_keys("barge")
    logger.info("Barge: USER ID 輸入完成")
    password_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Password' or contains(@id, 'mat-input-2')]")
    password_field.send_keys("123456")
    logger.info("Barge: PW 輸入完成")
    login_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'LOGIN') or contains(@class, 'mat-raised-button')]")))
    ActionChains(driver).move_to_element(login_button_barge).click().perform()
    logger.info("Barge: LOGIN 按鈕點擊成功")

# Barge 下載部分
@retry_on_failure(max_retries=MAX_RETRIES)
def process_barge_download(driver, wait, initial_files):
    logger.info("Barge: 直接前往 https://barge.oneport.com/downloadReport...")
    driver.get("https://barge.oneport.com/downloadReport")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    logger.info("Barge: downloadReport 頁面加載完成")
    start_time = time.time()
    report_type_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-form-field[.//mat-label[contains(text(), 'Report Type')]]//div[contains(@class, 'mat-select-trigger')]")))
    ActionChains(driver).move_to_element(report_type_trigger).click().perform()
    logger.info("Barge: Report Type 選擇開始")
    container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-option//span[contains(text(), 'Container Detail')]")))
    ActionChains(driver).move_to_element(container_detail_option).click().perform()
    logger.info("Barge: Container Detail 點擊成功")
    local_initial = initial_files.copy()
    download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Download']]")))
    ActionChains(driver).move_to_element(download_button_barge).click().perform()
    logger.info("Barge: Download 按鈕點擊成功")
    new_files = wait_for_new_file(local_initial, start_time, timeout=10, expected_pattern="ContainerDetailReport")
    filtered_files = {f for f in new_files if "ContainerDetailReport" in f}
    if len(filtered_files) >= BARGE_COUNT:
        logger.info(f"Barge: Container Detail 下載完成")
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
        new_files = process_barge_download(driver, wait, downloaded_files)
        downloaded_files.update(new_files)
        return downloaded_files
    finally:
        if driver:
            try:
                logout_toolbar_barge = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
                driver.execute_script("arguments[0].click();", logout_toolbar_barge)
                logout_button_barge = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Logout')]]")))
                driver.execute_script("arguments[0].click();", logout_button_barge)
                logger.info("Barge: 登出成功")
            except Exception as e:
                logger.error(f"Barge: 登出失敗: {str(e)}")
            driver.quit()
            logger.info("Barge WebDriver 關閉")

# 郵件發送函數（使用 tenacity 添加重試）
@tenacity.retry(stop=tenacity.stop_after_attempt(3), wait=tenacity.wait_fixed(5))
def send_email(downloaded_files):
    max_attachment_size = 25 * 1024 * 1024  # 25MB
    smtp_server = 'smtp.zoho.com'
    smtp_port = 587
    sender_email = os.environ.get('ZOHO_EMAIL', 'paklun_ckline@zohomail.com')
    sender_password = os.environ.get('ZOHO_PASSWORD', '@d6G.Pie5UkEPqm')
    receiver_email = 'paklun@ckline.com.hk'
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = f"[TESTING] HIT DAILY {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    body = "Attached are the daily reports downloaded from CPLUS (Container Movement Log, OnHand Container List, and Housekeeping Reports) and Barge (Container Detail)."
    msg.attach(MIMEText(body, 'plain'))
    for file in downloaded_files:
        file_path = os.path.join(download_dir, file)
        if os.path.getsize(file_path) > max_attachment_size:
            logger.error(f"文件 {file} 超過 25MB，無法附加")
            continue
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

# 主函數
def main():
    clear_download_dir()
    cplus_files = set()
    barge_files = set()
    house_file_count = [0]
    house_button_count = [0]

    def update_cplus_files_and_count(result):
        files, count, button_count = result
        cplus_files.update(files)
        house_file_count[0] = count
        house_button_count[0] = button_count

    with ThreadPoolExecutor(max_workers=2) as executor:
        cplus_future = executor.submit(process_cplus)
        barge_future = executor.submit(process_barge)
        update_cplus_files_and_count(cplus_future.result())
        barge_files.update(barge_future.result())

    logger.info("檢查所有下載文件...")
    downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
    expected_file_count = CPLUS_MOVEMENT_COUNT + CPLUS_ONHAND_COUNT + house_button_count[0] + BARGE_COUNT
    logger.info(f"預期文件數量: {expected_file_count}")
    if house_file_count[0] != house_button_count[0]:
        logger.error(f"Housekeeping Reports 下載文件數量（{house_file_count[0]}）不等於按鈕數量（{house_button_count[0]}），放棄發送郵件")
        return
    cplus_file_count = len([f for f in downloaded_files if "ContainerDetailReport" not in f])
    barge_file_count = len([f for f in downloaded_files if "ContainerDetailReport" in f])
    if cplus_file_count < (CPLUS_MOVEMENT_COUNT + CPLUS_ONHAND_COUNT + house_button_count[0]) or barge_file_count < BARGE_COUNT:
        logger.error(f"CPLUS 文件數量（{cplus_file_count}）或 Barge 文件數量（{barge_file_count}）不足")
        return
    if len(downloaded_files) == expected_file_count:
        logger.info(f"所有下載完成，檔案位於: {download_dir}")
        try:
            send_email(downloaded_files)
        except Exception as e:
            logger.error(f"郵件發送失敗: {str(e)}")
    else:
        logger.error(f"總下載文件數量不足（{len(downloaded_files)}/{expected_file_count}），放棄發送郵件")
    logger.info("腳本完成")

if __name__ == "__main__":
    setup_environment()
    main()
