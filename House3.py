# 確保無 pandas 導入
import os
import time
import shutil
import subprocess
from datetime import datetime
from selenium.webdriver.common.keys import Keys
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

cplus_download_dir = os.path.abspath("downloads_cplus")
barge_download_dir = os.path.abspath("downloads_barge")
MAX_RETRIES = 2
DOWNLOAD_TIMEOUT = 30  # 設置為 30 秒，避免下載超時

# 更新 get_chrome_options
def get_chrome_options(download_dir):
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('prefs', {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": False
    })
    chrome_options.binary_location = os.path.expanduser('~/chromium-bin/chromium-browser')
    return chrome_options

# 更新 setup_environment
def setup_environment():
    os.environ.pop('TZ', None)
    try:
        chromium_path = os.path.expanduser('~/chromium-bin/chromium-browser')
        chromedriver_path = os.path.expanduser('~/chromium-bin/chromedriver')
        if not os.path.exists(chromium_path) or not os.path.exists(chromedriver_path):
            subprocess.run(['sudo', 'apt-get', 'update', '-qq'], check=True)
            subprocess.run(['sudo', 'apt-get', 'install', '-y', 'chromium-browser', 'chromium-chromedriver'], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            os.makedirs(os.path.expanduser('~/chromium-bin'), exist_ok=True)
            subprocess.run(['sudo', 'cp', '/usr/bin/chromium-browser', chromium_path], check=True)
            subprocess.run(['sudo', 'cp', '/usr/bin/chromedriver', chromedriver_path], check=True)
            logging.info("Chromium 及 ChromeDriver 已安裝")
        else:
            logging.info("Chromium 及 ChromeDriver 已存在，跳過安裝")
            chromedriver_version = subprocess.run([chromedriver_path, '--version'], capture_output=True, text=True).stdout
            logging.info(f"Chromedriver 版本: {chromedriver_version}")

        required_packages = ['selenium', 'webdriver-manager', 'python-dotenv', 'pytz']
        for package in required_packages:
            result = subprocess.run(['pip', 'show', package], capture_output=True, text=True)
            if package not in result.stdout:
                logging.info(f"安裝 {package}...")
                subprocess.run(['pip', 'install', package], check=True)
            else:
                logging.info(f"{package} 已存在，跳過安裝")
    except subprocess.CalledProcessError as e:
        logging.error(f"環境準備失敗: {e}")
        raise

# 確保 check_env_vars 和 check_file 存在
def check_env_vars():
    required_vars = ['SITE_PASSWORD', 'BARGE_PASSWORD', 'ZOHO_EMAIL', 'ZOHO_PASSWORD', 'RECEIVER_EMAILS']
    missing = [var for var in required_vars if not os.environ.get(var)]
    if missing:
        logging.error(f"缺少環境變量: {', '.join(missing)}")
        raise EnvironmentError(f"缺少環境變量: {', '.join(missing)}")

def check_file(file_path):
    try:
        if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
            logging.info(f"檔案 {file_path} 存在且大小非零")
            return True
        else:
            logging.warning(f"檔案 {file_path} 不存在或大小為零")
            return False
    except Exception as e:
        logging.error(f"檢查檔案 {file_path} 失敗: {str(e)}")
        return False

def cleanup_downloads():
    for dir_path in [cplus_download_dir, barge_download_dir]:
        if os.path.exists(dir_path):
            shutil.rmtree(dir_path)
            os.makedirs(dir_path)
            logging.info(f"清理下載目錄: {dir_path}")

def get_chrome_options(download_dir):
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('prefs', {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": False
    })
    chrome_options.binary_location = os.path.expanduser('~/chromium-bin/chromium-browser')  # 更新路徑
    return chrome_options

def wait_for_new_file(download_dir, initial_files, timeout=DOWNLOAD_TIMEOUT):
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_files = set(os.listdir(download_dir))
        new_files = current_files - initial_files
        if new_files and any(os.path.getsize(os.path.join(download_dir, f)) > 0 for f in new_files):
            return new_files
        time.sleep(0.1)
    return set()

def handle_popup(driver, wait):
    max_attempts = 3
    attempt = 0
    while attempt < max_attempts:
        try:
            error_div = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'System Error') or contains(text(), 'Loading') or contains(@class, 'error') or contains(@class, 'popup')]")))
            logging.error(f"CPLUS/Barge: 檢測到系統錯誤: {error_div.text}")
            close_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Close') or contains(text(), 'OK') or contains(text(), 'Download Screenshots')]")))
            driver.execute_script("arguments[0].click();", close_button)
            logging.info("CPLUS/Barge: 關閉系統錯誤彈窗")
            time.sleep(1)
            driver.save_screenshot("system_error.png")
            with open("system_error.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            attempt += 1
        except TimeoutException:
            logging.debug("無系統錯誤或加載彈窗檢測到")
            break

def cplus_login(driver, wait):
    logging.info("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/")
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    logging.info(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}")
    time.sleep(0.5)

    logging.info("CPLUS: 點擊登錄前按鈕...")
    login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
    ActionChains(driver).move_to_element(login_button_pre).click().perform()
    logging.info("CPLUS: 登錄前按鈕點擊成功")
    time.sleep(0.5)

    logging.info("CPLUS: 輸入 COMPANY CODE...")
    company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    company_code_field.send_keys("CKL")
    logging.info("CPLUS: COMPANY CODE 輸入完成")
    time.sleep(0.5)

    logging.info("CPLUS: 輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
    user_id_field.send_keys("KEN")
    logging.info("CPLUS: USER ID 輸入完成")
    time.sleep(0.5)

    logging.info("CPLUS: 輸入 PASSWORD...")
    password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
    password_field.send_keys(os.environ.get('SITE_PASSWORD'))
    logging.info("CPLUS: PASSWORD 輸入完成")
    time.sleep(0.5)

    logging.info("CPLUS: 點擊 LOGIN 按鈕...")
    login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
    ActionChains(driver).move_to_element(login_button).click().perform()
    logging.info("CPLUS: LOGIN 按鈕點擊成功")
    try:
        handle_popup(driver, wait)
        WebDriverWait(driver, 60).until(EC.url_contains("https://cplus.hit.com.hk/app/#/"))
        logging.info("CPLUS: 檢測到主界面 URL，登錄成功")
    except TimeoutException:
        logging.warning("CPLUS: 登錄後主界面加載失敗，但繼續執行")
        driver.save_screenshot("login_warning.png")
        with open("login_warning.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)

def process_cplus_movement(driver, wait, initial_files):
    logging.info("CPLUS: 直接前往 Container Movement Log...")
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")), 60)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")), 60)
        logging.info("CPLUS: Container Movement Log 頁面加載完成")
    except TimeoutException:
        logging.error("CPLUS: Movement 頁面加載失敗，刷新頁面...")
        driver.refresh()
        time.sleep(5)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")), 60)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")), 60)
        logging.info("CPLUS: Container Movement Log 頁面加載完成 (刷新後)")

    logging.info("CPLUS: 等待表單字段加載...")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")), 60)
    logging.info("CPLUS: 表單字段加載完成")

    logging.info("CPLUS: 點擊 Search...")
    local_initial = initial_files.copy()
    try:
        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")), 60)
        ActionChains(driver).move_to_element(search_button).click().perform()
        logging.info("CPLUS: Search 按鈕點擊成功")
    except TimeoutException:
        logging.warning("CPLUS: Search 按鈕超時，刷新頁面...")
        driver.refresh()
        time.sleep(5)
        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")), 60)
        ActionChains(driver).move_to_element(search_button).click().perform()
        logging.info("CPLUS: Search 按鈕點擊成功 (刷新後)")

    logging.info("CPLUS: 點擊 Download...")
    for attempt in range(MAX_RETRIES):
        try:
            download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button")), 60)
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
            logging.error(f"CPLUS: Download 按鈕點擊失敗 (嘗試 {attempt+1}/{MAX_RETRIES}): {str(e)}")
            driver.save_screenshot(f"movement_download_failure_{attempt}.png")
            with open(f"movement_download_failure_{attempt}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            if attempt < MAX_RETRIES - 1:
                driver.refresh()
                time.sleep(5)
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
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")), 60)
        wait.until(EC.presence_of_element_located((By.XPATH, "//form//input")), 60)
        logging.info("CPLUS: OnHandContainerList 頁面加載完成")
    except TimeoutException:
        logging.error("CPLUS: OnHand 頁面加載失敗，刷新頁面...")
        driver.refresh()
        time.sleep(5)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")), 60)
        wait.until(EC.presence_of_element_located((By.XPATH, "//form//input")), 60)
        logging.info("CPLUS: OnHandContainerList 頁面加載完成 (刷新後)")

    logging.info("CPLUS: 等待表單字段加載...")
    wait.until(EC.presence_of_element_located((By.XPATH, "//form//input")), 60)
    logging.info("CPLUS: 表單字段加載完成")

    logging.info("CPLUS: 點擊 Search...")
    local_initial = initial_files.copy()
    try:
        search_button_onhand = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")), 60)
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
        time.sleep(5)
        search_button_onhand = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")), 60)
        try:
            ActionChains(driver).move_to_element(search_button_onhand).click().perform()
            logging.info("CPLUS: Search 按鈕點擊成功 (刷新後)")
        except ElementClickInterceptedException:
            logging.debug("CPLUS: Search 按鈕點擊被攔截，使用 JavaScript 點擊")
            driver.execute_script("arguments[0].click();", search_button_onhand)
            logging.info("CPLUS: Search 按鈕 JavaScript 點擊成功 (刷新後)")

    time.sleep(0.2)

    logging.info("CPLUS: 點擊 Export...")
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button")), 60)
    ActionChains(driver).move_to_element(export_button).click().perform()
    logging.info("CPLUS: Export 按鈕點擊成功")
    time.sleep(0.2)

    logging.info("CPLUS: 點擊 Export as CSV...")
    export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']")), 60)
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

    logging.info("CPLUS: 定位 Excel 下載按鈕...")
    local_initial = initial_files.copy()
    new_files = set()
    excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)]")
    button_count = len(excel_buttons)
    logging.info(f"CPLUS: 找到 {button_count} 個 Excel 下載按鈕")

    if button_count == 0:
        logging.debug("CPLUS: 未找到 Excel 按鈕，嘗試備用定位...")
        excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]//svg[@viewBox='0 0 24 24']//path[@fill='#036e11']")
        button_count = len(excel_buttons)
        logging.info(f"CPLUS: 備用定位找到 {button_count} 個 Excel 下載按鈕")

    for idx, button in enumerate(excel_buttons, 1):
        success = False
        try:
            # 確保按鈕可見同可點擊
            wait.until(EC.element_to_be_clickable(button))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
            
            # 獲取報告名稱
            try:
                report_name = driver.find_element(By.XPATH, f"//table[contains(@class, 'MuiTable-root')]//tbody//tr[{idx}]//td[3]").text
                logging.info(f"CPLUS: 準備點擊第 {idx} 個 Excel 按鈕，報告名稱: {report_name}")
            except:
                logging.debug(f"CPLUS: 無法獲取第 {idx} 個按鈕的報告名稱")

            # 用標準 click() 方法
            button.click()
            logging.info(f"CPLUS: 第 {idx} 個 Excel 下載按鈕標準點擊成功")
            
            # 處理彈出框
            handle_popup(driver, wait)
            
            # 等待新文件下載
            temp_new = wait_for_new_file(cplus_download_dir, local_initial)
            if temp_new:
                logging.info(f"CPLUS: 第 {idx} 個按鈕下載新文件: {', '.join(temp_new)}")
                local_initial.update(temp_new)
                new_files.update(temp_new)
                success = True
            else:
                logging.warning(f"CPLUS: 第 {idx} 個按鈕未觸發新文件下載")
                driver.save_screenshot(f"house_button_{idx}_failure.png")
                with open(f"house_button_{idx}_failure.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
        except ElementClickInterceptedException as e:
            logging.error(f"CPLUS: 第 {idx} 個按鈕點擊被阻擋: {str(e)}")
            driver.save_screenshot(f"house_button_{idx}_failure.png")
            with open(f"house_button_{idx}_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
        except Exception as e:
            logging.error(f"CPLUS: 第 {idx} 個按鈕點擊失敗: {str(e)}")
            driver.save_screenshot(f"house_button_{idx}_failure.png")
            with open(f"house_button_{idx}_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
        if not success:
            logging.warning(f"CPLUS: 第 {idx} 個 Excel 下載按鈕失敗")

    if new_files:
        logging.info(f"CPLUS: Housekeeping Reports 下載完成，共 {len(new_files)} 個文件，預期 {button_count} 個")
        return new_files, len(new_files), button_count
    else:
        logging.warning("CPLUS: Housekeeping Reports 未下載任何文件，記錄頁面狀態...")
        driver.save_screenshot("house_download_failure.png")
        with open("house_download_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: Housekeeping Reports 未下載任何文件")
        
def process_cplus_logout(driver, wait):
    try:
        logging.info("CPLUS: 嘗試登出...")
        logout_menu_button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", logout_menu_button)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", logout_menu_button)
        logging.info("CPLUS: 登錄按鈕點擊成功")
        try:
            logout_option = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='menu-list-grow']/div[6]/li")))
            driver.execute_script("arguments[0].scrollIntoView(true);", logout_option)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", logout_option)
            logging.info("CPLUS: Logout 選項點擊成功")
        except TimeoutException:
            logging.debug("CPLUS: 主 Logout 選項未找到，嘗試備用定位...")
            try:
                logout_option = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Logout')]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_option)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", logout_option)
                logging.info("CPLUS: 備用 Logout 選項 1 點擊成功")
            except TimeoutException:
                logging.debug("CPLUS: 備用 Logout 選項 1 失敗，嘗試備用定位 2...")
                try:
                    logout_option = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@role, 'menuitem') and contains(., 'Logout')]")))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_option)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", logout_option)
                    logging.info("CPLUS: 備用 Logout 選項 2 點擊成功")
                except TimeoutException:
                    logging.error("CPLUS: 所有 Logout 選項未找到，檢查頁面狀態...")
                    driver.save_screenshot("logout_failure.png")
                    with open("logout_failure.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                    raise Exception("CPLUS: Logout 失敗")
        time.sleep(2)
        try:
            login_button_pre = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")), 60)
            logging.info("CPLUS: Logout 成功，返回登錄頁")
        except TimeoutException:
            logging.error("CPLUS: Logout 失敗，未返回登錄頁")
            raise Exception("CPLUS: Logout 失敗")
    except Exception as logout_error:
        logging.error(f"CPLUS: 登出失敗: {str(logout_error)}")
        driver.save_screenshot("logout_failure.png")
        with open("logout_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise

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
        wait = WebDriverWait(driver, 60)
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
                        logging.info(f"CPLUS {section_name} 刷新頁面重試...")
                        driver.refresh()
                        time.sleep(5)
            if not success:
                logging.error(f"CPLUS {section_name} 經過 {MAX_RETRIES} 次嘗試失敗")
        process_cplus_logout(driver, wait)
        return downloaded_files, house_file_count, house_button_count, driver
    except Exception as e:
        logging.error(f"CPLUS 總錯誤: {str(e)}")
        return downloaded_files, house_file_count, house_button_count, driver
    finally:
        if driver:
            try:
                driver.quit()
                logging.info("CPLUS WebDriver 關閉")
            except Exception as e:
                logging.error(f"CPLUS WebDriver 關閉失敗: {str(e)}")

def barge_login(driver, wait):
    logging.info("Barge: 嘗試打開網站 https://barge.oneport.com/login...")
    driver.get("https://barge.oneport.com/login")
    logging.info(f"Barge: 網站已成功打開，當前 URL: {driver.current_url}")
    time.sleep(0.5)

    logging.info("Barge: 輸入 COMPANY ID...")
    company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Company ID' or contains(@id, 'mat-input-0')]")))
    company_id_field.send_keys("CKL")
    logging.info("Barge: COMPANY ID 輸入完成")
    time.sleep(0.5)

    logging.info("Barge: 輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='User ID' or contains(@id, 'mat-input-1')]")
    user_id_field.send_keys("barge")
    logging.info("Barge: USER ID 輸入完成")
    time.sleep(0.5)

    logging.info("Barge: 輸入 PW...")
    password_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Password' or contains(@id, 'mat-input-2')]")
    password_field.send_keys(os.environ.get('BARGE_PASSWORD', '123456'))
    logging.info("Barge: PW 輸入完成")
    time.sleep(0.5)

    logging.info("Barge: 點擊 LOGIN 按鈕...")
    login_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'LOGIN') or contains(@class, 'mat-raised-button')]")))
    ActionChains(driver).move_to_element(login_button_barge).click().perform()
    logging.info("Barge: LOGIN 按鈕點擊成功")
    time.sleep(0.5)
    
def process_barge_download(driver, wait, initial_files):
    logging.info("Barge: 直接前往 https://barge.oneport.com/downloadReport...")
    driver.get("https://barge.oneport.com/downloadReport")
    try:
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")), 60)
        logging.info("Barge: downloadReport 頁面加載完成")
    except TimeoutException:
        logging.error("Barge: downloadReport 頁面加載失敗，刷新頁面...")
        driver.refresh()
        time.sleep(5)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")), 60)
        logging.info("Barge: downloadReport 頁面加載完成 (刷新後)")

    logging.info("Barge: 選擇 Report Type...")
    report_type_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-form-field[.//mat-label[contains(text(), 'Report Type')]]//div[contains(@class, 'mat-select-trigger')]")), 60)
    ActionChains(driver).move_to_element(report_type_trigger).click().perform()
    logging.info("Barge: Report Type 選擇開始")
    time.sleep(0.5)

    logging.info("Barge: 點擊 Container Detail...")
    container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-option//span[contains(text(), 'Container Detail')]")), 60)
    ActionChains(driver).move_to_element(container_detail_option).click().perform()
    logging.info("Barge: Container Detail 點擊成功")
    time.sleep(0.5)

    logging.info("Barge: 點擊 Download...")
    local_initial = initial_files.copy()
    download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Download']]")), 60)
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

def process_barge_logout(driver, wait):
    try:
        logging.info("Barge: 嘗試登出...")
        logout_toolbar_barge = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar_barge)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", logout_toolbar_barge)
        logging.info("Barge: 工具欄點擊成功")
        try:
            logout_button_barge = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'mat-menu-panel')]//button//span[contains(text(), 'Logout')]")))
            driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", logout_button_barge)
            logging.info("Barge: Logout 選項點擊成功")
        except TimeoutException:
            logging.debug("Barge: 主 Logout 選項未找到，嘗試備用定位...")
            backup_logout_xpath = "//button[.//span[contains(text(), 'Logout')]]"
            logout_button_barge = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, backup_logout_xpath)))
            driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", logout_button_barge)
            logging.info("Barge: 備用 Logout 選項點擊成功")
        time.sleep(2)
        try:
            login_button_barge = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Company ID' or contains(@id, 'mat-input-0')]")), 60)
            logging.info("Barge: Logout 成功，返回登錄頁")
        except TimeoutException:
            logging.error("Barge: Logout 失敗，未返回登錄頁")
            raise Exception("Barge: Logout 失敗")
    except Exception as logout_error:
        logging.error(f"Barge: 登出失敗: {str(logout_error)}")
        driver.save_screenshot("barge_logout_failure.png")
        with open("barge_logout_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise

def send_email(subject, body, files, sender_email, receiver_emails, cc_emails, sender_password, dry_run=False):
    msg = MIMEMultipart('alternative')
    msg['From'] = sender_email
    msg['To'] = ', '.join(receiver_emails)
    if cc_emails:
        msg['Cc'] = ', '.join(cc_emails)
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'html'))
    plain_text = body.replace('<br>', '\n').replace('<table>', '').replace('</table>', '').replace('<tr>', '\n').replace('<td>', ' | ').replace('</td>', '').replace('<th>', ' | ').replace('</th>', '').strip()
    msg.attach(MIMEText(plain_text, 'plain'))

    for file in files:
        if file in os.listdir(cplus_download_dir):
            file_path = os.path.join(cplus_download_dir, file)
        else:
            file_path = os.path.join(barge_download_dir, file)
        if os.path.exists(file_path):
            attachment = MIMEBase('application', 'octet-stream')
            with open(file_path, 'rb') as f:
                attachment.set_payload(f.read())
            encoders.encode_base64(attachment)
            attachment.add_header('Content-Disposition', f'attachment; filename={file}')
            msg.attach(attachment)
        else:
            logging.warning(f"附件不存在: {file_path}")

    if not dry_run:
        try:
            smtp_server = os.environ.get('SMTP_SERVER', 'smtp.zoho.com')
            smtp_port = int(os.environ.get('SMTP_PORT', 587))
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                all_receivers = receiver_emails + cc_emails
                server.sendmail(sender_email, all_receivers, msg.as_string())
            logging.info(f"郵件發送成功至 {', '.join(all_receivers)}")
        except smtplib.SMTPAuthenticationError:
            logging.error("SMTP 認證失敗：檢查用戶名/密碼")
        except smtplib.SMTPConnectError:
            logging.error("SMTP 連接失敗：檢查伺服器/端口")
        except Exception as e:
            logging.error(f"郵件發送失敗: {str(e)}")
    else:
        logging.info(f"模擬發送郵件：\nFrom: {sender_email}\nTo: {msg['To']}\nCc: {msg.get('Cc', '')}\nSubject: {msg['Subject']}\nBody: {body}")

def process_barge():
    driver = None
    downloaded_files = set()
    initial_files = set(os.listdir(barge_download_dir))
    try:
        driver = webdriver.Chrome(options=get_chrome_options(barge_download_dir))
        logging.info("Barge WebDriver 初始化成功")
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        wait = WebDriverWait(driver, 60)
        barge_login(driver, wait)
        new_files = process_barge_download(driver, wait, initial_files)
        downloaded_files.update(new_files)
        initial_files.update(new_files)
        process_barge_logout(driver, wait)
        return downloaded_files, driver
    except Exception as e:
        logging.error(f"Barge 總錯誤: {str(e)}")
        return downloaded_files, driver
    finally:
        if driver:
            try:
                driver.quit()
                logging.info("Barge WebDriver 關閉")
            except Exception as e:
                logging.error(f"Barge WebDriver 關閉失敗: {str(e)}")

def main():
    load_dotenv()
    logging.info("開始執行 House3.py")
    check_env_vars()  # 檢查環境變量
    clear_download_dirs()
    logging.info("下載目錄已清理")

    # 順序執行 CPLUS 和 Barge
    cplus_files, cplus_house_file_count, cplus_house_button_count, cplus_driver = set(), 0, 0, None
    barge_files, barge_driver = set(), None

    try:
        logging.info("開始處理 CPLUS")
        cplus_files, cplus_house_file_count, cplus_house_button_count, cplus_driver = process_cplus()
        logging.info("CPLUS 處理完成")
    except Exception as e:
        logging.error(f"CPLUS 處理失敗: {str(e)}")
    finally:
        if cplus_driver:
            try:
                cplus_driver.quit()
                logging.info("CPLUS WebDriver 關閉")
            except Exception as e:
                logging.error(f"CPLUS WebDriver 關閉失敗: {str(e)}")

    try:
        logging.info("開始處理 Barge")
        barge_files, barge_driver = process_barge()
        logging.info("Barge 處理完成")
    except Exception as e:
        logging.error(f"Barge 處理失敗: {str(e)}")
    finally:
        if barge_driver:
            try:
                barge_driver.quit()
                logging.info("Barge WebDriver 關閉")
            except Exception as e:
                logging.error(f"Barge WebDriver 關閉失敗: {str(e)}")

    downloaded_files = cplus_files.union(barge_files)
    logging.info(f"總下載文件: {len(downloaded_files)} 個")
    for file in downloaded_files:
        logging.info(f"找到檔案: {file}")

    # 檢查檔案
    for file in downloaded_files:
        file_path = os.path.join(cplus_download_dir if file in cplus_files else barge_download_dir, file)
        if not check_file(file_path):
            logging.warning(f"檔案 {file} 檢查失敗")

    # 檢查必須檔案
    required_patterns = {'movement': 'cntrMoveLog', 'onhand': 'data_', 'barge': 'ContainerDetailReport'}
    housekeep_prefixes = ['IA5_', 'DM1C_', 'IA15_', 'GA1_', 'IA15_', 'IA17_']
    has_required = all(any(pattern in f for f in downloaded_files) for pattern in required_patterns.values())
    house_files = [f for f in downloaded_files if any(p in f for p in housekeep_prefixes)]
    house_download_count = len(set(house_files))
    house_ok = (house_download_count >= cplus_house_button_count - 1) or (cplus_house_button_count == 0)

    if has_required and house_ok:
        logging.info("所有必須文件齊全，準備發送郵件")
        try:
            smtp_server = os.environ.get('SMTP_SERVER', 'smtp.zoho.com')
            smtp_port = int(os.environ.get('SMTP_PORT', 587))
            sender_email = os.environ['ZOHO_EMAIL']
            sender_password = os.environ['ZOHO_PASSWORD']
            receiver_emails = os.environ.get('RECEIVER_EMAILS', 'paklun@ckline.com.hk').split(',')
            cc_emails = os.environ.get('CC_EMAILS', '').split(',') if os.environ.get('CC_EMAILS') else []
            dry_run = os.environ.get('DRY_RUN', 'False').lower() == 'true'

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
            send_email(
                f"[TESTING] HIT DAILY {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                body_html,
                downloaded_files,
                sender_email,
                receiver_emails,
                cc_emails,
                sender_password,
                dry_run
            )
            logging.info("郵件發送完成")
        except Exception as e:
            logging.error(f"郵件發送失敗: {str(e)}")
    else:
        logging.warning(f"文件不齊全: 缺少必須文件 (has_required={has_required}) 或 House文件不足 (download={house_download_count}, button={cplus_house_button_count})")

    cleanup_downloads()
    logging.info("腳本執行完成")
