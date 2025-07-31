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
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# 配置日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('script.log')
    ]
)
logger = logging.getLogger(__name__)

# 全局變量
download_dir = os.path.abspath("downloads")
cplus_dir = os.path.join(download_dir, "cplus")
barge_dir = os.path.join(download_dir, "barge")
EXPECTED_CPLUS_COUNT = 8  # 1 (Container Movement Log) + 1 (OnHandContainerList) + 6 (Housekeeping Reports)
EXPECTED_BARGE_COUNT = 1  # 1 (Barge)
EXPECTED_TOTAL_COUNT = EXPECTED_CPLUS_COUNT + EXPECTED_BARGE_COUNT

# 清空特定下載目錄
def clear_specific_dir(dir_path):
    if os.path.exists(dir_path):
        shutil.rmtree(dir_path)
    os.makedirs(dir_path)
    logger.info(f"清空並創建下載目錄: {dir_path}")

# 確保環境準備
def setup_environment():
    try:
        result = subprocess.run(['which', 'chromium-browser'], capture_output=True, text=True)
        if result.returncode != 0:
            logger.info("正在安裝 Chromium 及 ChromeDriver...")
            subprocess.run(['sudo', 'apt-get', 'update', '-qq'], check=True)
            subprocess.run(['sudo', 'apt-get', 'install', '-y', 'chromium-browser', 'chromium-chromedriver'], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            logger.info("Chromium 及 ChromeDriver 已安裝")
        else:
            logger.info("Chromium 及 ChromeDriver 已存在，跳過安裝")

        result = subprocess.run(['pip', 'show', 'selenium'], capture_output=True, text=True)
        if "selenium" not in result.stdout or "webdriver-manager" not in subprocess.run(['pip', 'show', 'webdriver-manager'], capture_output=True, text=True).stdout:
            logger.info("正在安裝 Selenium 及 WebDriver Manager...")
            subprocess.run(['pip', 'install', 'selenium', 'webdriver-manager'], check=True)
            logger.info("Selenium 及 WebDriver Manager 已安裝")
        else:
            logger.info("Selenium 及 WebDriver Manager 已存在，跳過安裝")
    except subprocess.CalledProcessError as e:
        logger.error(f"環境準備失敗: {e}")
        raise

# 設置 Chrome 選項
def get_chrome_options(download_path):
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
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": False
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.binary_location = '/usr/bin/chromium-browser'
    return chrome_options

# 檢查新文件出現並驗證文件完整性
def wait_for_new_file(dir_path, initial_files, timeout=15):
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_files = set(f for f in os.listdir(dir_path) if f.endswith(('.csv', '.xlsx')))
        new_files = current_files - initial_files
        for file in new_files:
            file_path = os.path.join(dir_path, file)
            if os.path.getsize(file_path) > 0:  # 確保文件不為空
                logger.info(f"檢測到新文件: {file}")
                return new_files
        time.sleep(0.1)
    logger.warning("未檢測到新文件")
    return set()

# CPLUS 操作
def process_cplus():
    dir_path = cplus_dir
    driver = None
    downloaded_files = set()
    try:
        options = get_chrome_options(dir_path)
        driver = webdriver.Chrome(options=options)
        logger.info("CPLUS WebDriver 初始化成功")
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

        # 前往登入頁面 (CPLUS)
        logger.info("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/")
        driver.get("https://cplus.hit.com.hk/frontpage/#/")
        logger.info(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}")
        time.sleep(2)

        # 點擊登錄前按鈕
        logger.info("CPLUS: 點擊登錄前按鈕...")
        wait = WebDriverWait(driver, 20)
        login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
        ActionChains(driver).move_to_element(login_button_pre).click().perform()
        logger.info("CPLUS: 登錄前按鈕點擊成功")
        time.sleep(2)

        # 輸入 COMPANY CODE
        logger.info("CPLUS: 輸入 COMPANY CODE...")
        company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
        company_code_field.send_keys(os.environ.get('CPLUS_COMPANY_CODE', 'CKL'))
        logger.info("CPLUS: COMPANY CODE 輸入完成")
        time.sleep(1)

        # 輸入 USER ID
        logger.info("CPLUS: 輸入 USER ID...")
        user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
        user_id_field.send_keys(os.environ.get('CPLUS_USER_ID', 'KEN'))
        logger.info("CPLUS: USER ID 輸入完成")
        time.sleep(1)

        # 輸入 PASSWORD
        logger.info("CPLUS: 輸入 PASSWORD...")
        password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
        password = os.environ.get('CPLUS_PASSWORD')
        if not password:
            logger.error("CPLUS: 未設置環境變量 CPLUS_PASSWORD")
            raise ValueError("CPLUS_PASSWORD 環境變量未設置")
        password_field.send_keys(password)
        logger.info("CPLUS: PASSWORD 輸入完成")
        time.sleep(1)

        # 點擊 LOGIN 按鈕
        logger.info("CPLUS: 點擊 LOGIN 按鈕...")
        login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
        ActionChains(driver).move_to_element(login_button).click().perform()
        logger.info("CPLUS: LOGIN 按鈕點擊成功")
        time.sleep(2)

        # 前往 Container Movement Log 頁面
        logger.info("CPLUS: 直接前往 Container Movement Log...")
        driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
        time.sleep(2)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")))
        logger.info("CPLUS: Container Movement Log 頁面加載完成")

        # 點擊 Search
        logger.info("CPLUS: 點擊 Search...")
        initial_files = set(f for f in os.listdir(dir_path) if f.endswith(('.csv', '.xlsx')))
        for attempt in range(2):
            try:
                search_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")))
                ActionChains(driver).move_to_element(search_button).click().perform()
                logger.info("CPLUS: Search 按鈕點擊成功")
                break
            except TimeoutException:
                logger.warning(f"CPLUS: Search 按鈕未找到，嘗試備用定位 {attempt+1}/2...")
                try:
                    search_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and .//span[contains(text(), 'Search')]]")))
                    ActionChains(driver).move_to_element(search_button).click().perform()
                    logger.info("CPLUS: 備用 Search 按鈕 1 點擊成功")
                    break
                except TimeoutException:
                    logger.warning(f"CPLUS: 備用 Search 按鈕 1 失敗，嘗試備用定位 2 (嘗試 {attempt+1}/2)...")
                    try:
                        search_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]")))
                        ActionChains(driver).move_to_element(search_button).click().perform()
                        logger.info("CPLUS: 備用 Search 按鈕 2 點擊成功")
                        break
                    except TimeoutException:
                        logger.error(f"CPLUS: 備用 Search 按鈕 2 失敗 (嘗試 {attempt+1}/2)")
        else:
            logger.error("CPLUS: Container Movement Log Search 按鈕點擊失敗，重試 2 次後放棄")

        # 點擊 Download
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
            logger.error("CPLUS: Container Movement Log Download 按鈕點擊失敗，重試 2 次後放棄")

        # 檢查新文件
        new_files = wait_for_new_file(dir_path, initial_files, timeout=15)
        if new_files:
            logger.info(f"CPLUS: Container Movement Log 下載完成，檔案位於: {dir_path}")
            for file in new_files:
                logger.info(f"CPLUS: 新下載檔案: {file}")
            downloaded_files.update(new_files)
        else:
            logger.warning("CPLUS: Container Movement Log 未觸發新文件下載")

        # 前往 OnHandContainerList 頁面
        logger.info("CPLUS: 前往 OnHandContainerList 頁面...")
        driver.get("https://cplus.hit.com.hk/app/#/enquiry/OnHandContainerList")
        time.sleep(2)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
        logger.info("CPLUS: OnHandContainerList 頁面加載完成")

        # 點擊 Search
        logger.info("CPLUS: 點擊 Search...")
        try:
            search_button_onhand = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")))
            ActionChains(driver).move_to_element(search_button_onhand).click().perform()
            logger.info("CPLUS: Search 按鈕點擊成功")
        except TimeoutException:
            logger.warning("CPLUS: Search 按鈕未找到，嘗試備用定位...")
            search_button_onhand = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]")))
            ActionChains(driver).move_to_element(search_button_onhand).click().perform()
            logger.info("CPLUS: 備用 Search 按鈕點擊成功")
        time.sleep(0.5)

        # 點擊 Export
        logger.info("CPLUS: 點擊 Export...")
        export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button")))
        ActionChains(driver).move_to_element(export_button).click().perform()
        logger.info("CPLUS: Export 按鈕點擊成功")
        time.sleep(0.5)

        # 點擊 Export as CSV
        logger.info("CPLUS: 點擊 Export as CSV...")
        export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']")))
        ActionChains(driver).move_to_element(export_csv_button).click().perform()
        logger.info("CPLUS: Export as CSV 按鈕點擊成功")
        time.sleep(0.5)

        # 檢查新文件
        new_files = wait_for_new_file(dir_path, initial_files, timeout=5)
        if new_files:
            logger.info(f"CPLUS: OnHandContainerList 下載完成，檔案位於: {dir_path}")
            for file in new_files:
                logger.info(f"CPLUS: 新下載檔案: {file}")
            downloaded_files.update(new_files)
            initial_files.update(new_files)
        else:
            logger.warning("CPLUS: OnHandContainerList 未觸發新文件下載")

        # 前往 Housekeeping Reports 頁面
        logger.info("CPLUS: 前往 Housekeeping Reports 頁面...")
        driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
        time.sleep(2)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
        logger.info("CPLUS: Housekeeping Reports 頁面加載完成")

        # 等待表格加載完成
        logger.info("CPLUS: 等待表格加載...")
        try:
            wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
            logger.info("CPLUS: 表格加載完成")
        except TimeoutException:
            logger.warning("CPLUS: 表格未加載，跳過下載步驟")
            return downloaded_files

        # 定位並點擊所有 Excel 下載按鈕
        logger.info("CPLUS: 定位並點擊所有 Excel 下載按鈕...")
        initial_files_house = downloaded_files.copy()
        excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]//svg[@viewBox='0 0 24 24']//path[@fill='#036e11']")
        button_count = len(excel_buttons)
        logger.info(f"CPLUS: 找到 {button_count} 個 Excel 下載按鈕")

        # 備用定位
        if button_count == 0:
            logger.warning("CPLUS: 未找到 Excel 按鈕，嘗試備用定位...")
            excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)]")
            button_count = len(excel_buttons)
            logger.info(f"CPLUS: 備用定位找到 {button_count} 個 Excel 下載按鈕")

        for idx, button in enumerate(excel_buttons):
            try:
                # 確保按鈕可點擊
                button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"(//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)])[{idx+1}]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", button)
                time.sleep(0.5)

                # 記錄報告名稱
                try:
                    report_name = driver.find_element(By.XPATH, f"//table[contains(@class, 'MuiTable-root')]//tbody//tr[{idx+1}]//td[3]").text
                    logger.info(f"CPLUS: 準備點擊第 {idx+1} 個 Excel 按鈕，報告名稱: {report_name}")
                except:
                    logger.warning(f"CPLUS: 無法獲取第 {idx+1} 個按鈕的報告名稱")

                # ActionChains 點擊（重試 2 次）
                for attempt in range(2):
                    try:
                        ActionChains(driver).move_to_element(button).pause(0.5).click().perform()
                        logger.info(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕點擊成功")
                        break
                    except Exception as e:
                        logger.warning(f"CPLUS: 第 {idx+1} 個按鈕點擊嘗試 {attempt+1} 失敗: {str(e)}")
                        time.sleep(0.5)
                else:
                    logger.error(f"CPLUS: 第 {idx+1} 個按鈕點擊失敗，重試 2 次後放棄")
                    continue

                # 檢查新文件
                new_files = wait_for_new_file(dir_path, initial_files_house, timeout=15)
                if new_files:
                    logger.info(f"CPLUS: 第 {idx+1} 個按鈕下載新文件: {', '.join(new_files)}")
                    initial_files_house.update(new_files)
                    downloaded_files.update(new_files)
                else:
                    logger.warning(f"CPLUS: 第 {idx+1} 個按鈕未觸發新文件下載")

                # 備用 JavaScript 點擊
                try:
                    driver.execute_script("arguments[0].click();", button)
                    logger.info(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕 JavaScript 點擊成功")
                    new_files = wait_for_new_file(dir_path, initial_files_house, timeout=15)
                    if new_files:
                        logger.info(f"CPLUS: 第 {idx+1} 個按鈕下載新文件 (JavaScript): {', '.join(new_files)}")
                        initial_files_house.update(new_files)
                        downloaded_files.update(new_files)
                    else:
                        logger.warning(f"CPLUS: 第 {idx+1} 個按鈕 JavaScript 未觸發新文件下載")
                except Exception as js_e:
                    logger.warning(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕 JavaScript 點擊失敗: {str(js_e)}")

            except ElementClickInterceptedException as e:
                logger.error(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕點擊被攔截: {str(e)}")
            except TimeoutException as e:
                logger.error(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕不可點擊: {str(e)}")
            except Exception as e:
                logger.error(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕點擊失敗: {str(e)}")

        return downloaded_files

    except Exception as e:
        logger.error(f"CPLUS 錯誤: {str(e)}")
        return downloaded_files

    finally:
        # 確保登出
        try:
            if driver:
                logger.info("CPLUS: 嘗試登出...")
                logout_menu_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_menu_button)
                time.sleep(0.5)
                ActionChains(driver).move_to_element(logout_menu_button).click().perform()
                logger.info("CPLUS: 登錄按鈕點擊成功")

                logout_option = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='menu-list-grow']/div[6]/li")))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_option)
                time.sleep(0.5)
                ActionChains(driver).move_to_element(logout_option).click().perform()
                logger.info("CPLUS: Logout 選項點擊成功")
                time.sleep(2)
        except TimeoutException:
            logger.warning("CPLUS: 登出按鈕未找到，嘗試備用定位...")
            try:
                logout_option = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Logout')]")))
                ActionChains(driver).move_to_element(logout_option).click().perform()
                logger.info("CPLUS: 備用 Logout 選項點擊成功")
                time.sleep(2)
            except TimeoutException:
                logger.warning("CPLUS: 備用 Logout 選項未找到，跳過登出")
        except Exception as logout_error:
            logger.error(f"CPLUS: 登出失敗: {str(logout_error)}")

        if driver:
            driver.quit()
            logger.info("CPLUS WebDriver 關閉")

# Barge 操作
def process_barge():
    dir_path = barge_dir
    driver = None
    downloaded_files = set()
    try:
        options = get_chrome_options(dir_path)
        driver = webdriver.Chrome(options=options)
        logger.info("Barge WebDriver 初始化成功")
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

        # 前往登入頁面
        logger.info("Barge: 嘗試打開網站 https://barge.oneport.com/bargeBooking...")
        driver.get("https://barge.oneport.com/bargeBooking")
        logger.info(f"Barge: 網站已成功打開，當前 URL: {driver.current_url}")
        time.sleep(3)

        # 輸入 COMPANY ID
        logger.info("Barge: 輸入 COMPANY ID...")
        wait = WebDriverWait(driver, 20)
        company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='mat-input-0']")))
        company_id_field.send_keys(os.environ.get('BARGE_COMPANY_ID', 'CKL'))
        logger.info("Barge: COMPANY ID 輸入完成")
        time.sleep(1)

        # 輸入 USER ID
        logger.info("Barge: 輸入 USER ID...")
        user_id_field_barge = driver.find_element(By.XPATH, "//*[@id='mat-input-1']")
        user_id_field_barge.send_keys(os.environ.get('BARGE_USER_ID', 'barge'))
        logger.info("Barge: USER ID 輸入完成")
        time.sleep(1)

        # 輸入 PW
        logger.info("Barge: 輸入 PW...")
        password_field_barge = driver.find_element(By.XPATH, "//*[@id='mat-input-2']")
        password = os.environ.get('BARGE_PASSWORD')
        if not password:
            logger.error("Barge: 未設置環境變量 BARGE_PASSWORD")
            raise ValueError("BARGE_PASSWORD 環境變量未設置")
        password_field_barge.send_keys(password)
        logger.info("Barge: PW 輸入完成")
        time.sleep(1)

        # 點擊 LOGIN
        logger.info("Barge: 點擊 LOGIN 按鈕...")
        login_button_barge = driver.find_element(By.XPATH, "//*[@id='login-form-container']/app-login-form/form/div/button")
        ActionChains(driver).move_to_element(login_button_barge).click().perform()
        logger.info("Barge: LOGIN 按鈕點擊成功")
        time.sleep(3)

        # 直接前往 downloadReport 頁面
        logger.info("Barge: 直接前往 https://barge.oneport.com/downloadReport...")
        driver.get("https://barge.oneport.com/downloadReport")
        time.sleep(3)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='content-mount']")))
        logger.info("Barge: downloadReport 頁面加載完成")

        # 選擇 Report Type
        logger.info("Barge: 選擇 Report Type...")
        report_type_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-select-value-61']/span")))
        ActionChains(driver).move_to_element(report_type_select).click().perform()
        logger.info("Barge: Report Type 選擇開始")
        time.sleep(2)

        # 點擊 Container Detail
        logger.info("Barge: 點擊 Container Detail...")
        container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-option-508']/span")))
        ActionChains(driver).move_to_element(container_detail_option).click().perform()
        logger.info("Barge: Container Detail 點擊成功")
        time.sleep(2)

        # 點擊 Download
        logger.info("Barge: 點擊 Download...")
        initial_files = set(f for f in os.listdir(dir_path) if f.endswith(('.csv', '.xlsx')))
        for attempt in range(2):
            try:
                download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='content-mount']/app-download-report/div[2]/div/form/div[2]/button")))
                ActionChains(driver).move_to_element(download_button_barge).click().perform()
                logger.info("Barge: Download 按鈕點擊成功")
                time.sleep(0.5)
                break
            except Exception as e:
                logger.warning(f"Barge: Download 按鈕點擊失敗 (嘗試 {attempt+1}/2): {str(e)}")
                time.sleep(0.5)
        else:
            logger.error("Barge: Container Detail Download 按鈕點擊失敗，重試 2 次後放棄")

        # 檢查新文件
        new_files = wait_for_new_file(dir_path, initial_files, timeout=15)
        if new_files:
            logger.info(f"Barge: Container Detail 下載完成，檔案位於: {dir_path}")
            for file in new_files:
                logger.info(f"Barge: 新下載檔案: {file}")
            downloaded_files.update(new_files)
        else:
            logger.warning("Barge: Container Detail 未觸發新文件下載")

        return downloaded_files

    except Exception as e:
        logger.error(f"Barge 錯誤: {str(e)}")
        return downloaded_files

    finally:
        # 確保登出
        try:
            if driver:
                logger.info("Barge: 點擊工具欄進行登出...")
                logout_toolbar_barge = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar_barge)
                time.sleep(0.5)
                ActionChains(driver).move_to_element(logout_toolbar_barge).click().perform()
                logger.info("Barge: 工具欄點擊成功")

                logger.info("Barge: 點擊 Logout 選項...")
                logout_button_barge = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-menu-panel-11']/div/button/span")))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                time.sleep(0.5)
                ActionChains(driver).move_to_element(logout_button_barge).click().perform()
                logger.info("Barge: Logout 選項點擊成功")
                time.sleep(2)
        except TimeoutException:
            logger.warning("Barge: 登出按鈕未找到，嘗試備用定位...")
            try:
                logout_button_barge = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Logout')]")))
                ActionChains(driver).move_to_element(logout_button_barge).click().perform()
                logger.info("Barge: 備用 Logout 選項點擊成功")
                time.sleep(2)
            except TimeoutException:
                logger.warning("Barge: 備用 Logout 選項未找到，跳過登出")
        except Exception as e:
            logger.error(f"Barge: 登出失敗: {str(e)}")

        if driver:
            driver.quit()
            logger.info("Barge WebDriver 關閉")

# 運行過程並重試
def run_process_with_retry(process_func, dir_path, expected_count, max_retries=3):
    downloaded_files = set()
    for attempt in range(max_retries):
        clear_specific_dir(dir_path)
        new_files = process_func()
        downloaded_files = new_files  # 既然清空了，每次都是新的
        current_count = len([f for f in os.listdir(dir_path) if f.endswith(('.csv', '.xlsx'))])
        if current_count >= expected_count:
            logger.info(f"{process_func.__name__} 下載成功，文件數: {current_count}")
            break
        else:
            logger.warning(f"{process_func.__name__} 文件數不足 ({current_count}/{expected_count})，準備第 {attempt + 2} 次重試...")
            time.sleep(5)
    else:
        logger.error(f"{process_func.__name__} 已達最大重試次數 ({max_retries})，文件數: {len(downloaded_files)}")
    return downloaded_files

# 主函數
def main():
    clear_specific_dir(download_dir)  # 清空最終目錄

    # 運行 CPLUS 過程
    cplus_files = run_process_with_retry(process_cplus, cplus_dir, EXPECTED_CPLUS_COUNT)

    # 運行 Barge 過程
    barge_files = run_process_with_retry(process_barge, barge_dir, EXPECTED_BARGE_COUNT)

    # 複製文件到最終目錄
    for file in cplus_files:
        shutil.copy(os.path.join(cplus_dir, file), download_dir)
    for file in barge_files:
        shutil.copy(os.path.join(barge_dir, file), download_dir)

    # 檢查所有下載文件
    logger.info("檢查所有下載文件...")
    downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
    total_count = len(downloaded_files)
    if total_count >= EXPECTED_TOTAL_COUNT:
        logger.info(f"所有下載完成，檔案位於: {download_dir}，總文件數: {total_count}")
        for file in downloaded_files:
            logger.info(f"找到檔案: {file}")

        # 發送 Zoho Mail
        logger.info("開始發送郵件...")
        try:
            smtp_server = 'smtp.zoho.com'
            smtp_port = 587
            sender_email = os.environ.get('ZOHO_EMAIL', 'paklun_ckline@zohomail.com')
            sender_password = os.environ.get('ZOHO_PASSWORD')
            if not sender_password:
                logger.error("未設置環境變量 ZOHO_PASSWORD")
                raise ValueError("ZOHO_PASSWORD 環境變量未設置")
            receiver_email = 'paklun@ckline.com.hk'

            # 創建郵件
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = receiver_email
            msg['Subject'] = f"[TESTING] HIT DAILY {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

            # 新增郵件正文
            body = "Attached are the daily reports downloaded from CPLUS (Container Movement Log, OnHand Container List, and Housekeeping Reports) and Barge (Container Detail)."
            msg.attach(MIMEText(body, 'plain'))

            # 添加附件
            for file in downloaded_files:
                file_path = os.path.join(download_dir, file)
                attachment = MIMEBase('application', 'octet-stream')
                attachment.set_payload(open(file_path, 'rb').read())
                encoders.encode_base64(attachment)
                attachment.add_header('Content-Disposition', f'attachment; filename={file}')
                msg.attach(attachment)

            # 連接 SMTP 伺服器並發送
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
            server.quit()
            logger.info("郵件發送成功!")
        except Exception as e:
            logger.error(f"郵件發送失敗: {str(e)}")
    else:
        logger.error(f"下載文件數量不足（{total_count}/{EXPECTED_TOTAL_COUNT}），無法發送郵件")

    logger.info("腳本完成")

if __name__ == "__main__":
    setup_environment()
    main()
