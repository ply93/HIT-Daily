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
# 全局變量
download_dir = os.path.abspath("downloads")
CPLUS_MOVEMENT_COUNT = 1 # Container Movement Log
CPLUS_ONHAND_COUNT = 1 # OnHandContainerList
BARGE_COUNT = 1 # Barge
MAX_RETRIES = 3
# 清空下載目錄
def clear_download_dir():
    if os.path.exists(download_dir):
        shutil.rmtree(download_dir)
    os.makedirs(download_dir)
    print(f"創建下載目錄: {download_dir}", flush=True)
# 確保環境準備
def setup_environment():
    try:
        result = subprocess.run(['which', 'chromium-browser'], capture_output=True, text=True)
        if result.returncode != 0:
            subprocess.run(['sudo', 'apt-get', 'update', '-qq'], check=True)
            subprocess.run(['sudo', 'apt-get', 'install', '-y', 'chromium-browser', 'chromium-chromedriver'], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            print("Chromium 及 ChromeDriver 已安裝", flush=True)
        else:
            print("Chromium 及 ChromeDriver 已存在，跳過安裝", flush=True)
        result = subprocess.run(['pip', 'show', 'selenium'], capture_output=True, text=True)
        if "selenium" not in result.stdout or "webdriver-manager" not in subprocess.run(['pip', 'show', 'webdriver-manager'], capture_output=True, text=True).stdout:
            subprocess.run(['pip', 'install', 'selenium', 'webdriver-manager'], check=True)
            print("Selenium 及 WebDriver Manager 已安裝", flush=True)
        else:
            print("Selenium 及 WebDriver Manager 已存在，跳過安裝", flush=True)
    except subprocess.CalledProcessError as e:
        print(f"環境準備失敗: {e}", flush=True)
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
# 檢查新文件出現 (動態超時)
def wait_for_new_file(initial_files, start_time, timeout=10, expected_pattern=None):
    max_timeout = max(10, timeout * 1.5)  # 動態增加 50% 緩衝
    end_time = time.time() + max_timeout
    while time.time() < end_time:
        current_files = set(f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx')))
        new_files = current_files - initial_files
        if new_files:
            # 過濾文件：只保留在 start_time 之後創建且匹配預期模式的文件
            filtered_files = set()
            for file in new_files:
                file_path = os.path.join(download_dir, file)
                if os.path.getmtime(file_path) >= start_time:
                    if expected_pattern is None or expected_pattern in file:
                        filtered_files.add(file)
            if filtered_files:
                return filtered_files
        time.sleep(0.5)
    return set()
# 等待 AJAX 加載完成
def wait_for_ajax(driver, timeout=10):
    wait = WebDriverWait(driver, timeout)
    wait.until(lambda d: d.execute_script("return (typeof jQuery === 'undefined' || jQuery.active == 0) && document.readyState === 'complete'"))
# CPLUS 登入
def cplus_login(driver, wait):
    print("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/", flush=True)
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    print(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)
    wait_for_ajax(driver)
    print("CPLUS: 點擊登錄前按鈕...", flush=True)
    login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Login']]")))  # 優化為更穩定選擇器
    ActionChains(driver).move_to_element(login_button_pre).click().perform()
    print("CPLUS: 登錄前按鈕點擊成功", flush=True)
    wait_for_ajax(driver)
    print("CPLUS: 輸入 COMPANY CODE...", flush=True)
    company_code_field = wait.until(EC.presence_of_element_located((By.ID, "companyCode")))
    company_code_field.send_keys("CKL")
    print("CPLUS: COMPANY CODE 輸入完成", flush=True)
    print("CPLUS: 輸入 USER ID...", flush=True)
    user_id_field = wait.until(EC.presence_of_element_located((By.ID, "userId")))
    user_id_field.send_keys("KEN")
    print("CPLUS: USER ID 輸入完成", flush=True)
    print("CPLUS: 輸入 PASSWORD...", flush=True)
    password_field = wait.until(EC.presence_of_element_located((By.ID, "passwd")))
    password_field.send_keys(os.environ.get('SITE_PASSWORD'))
    print("CPLUS: PASSWORD 輸入完成", flush=True)
    print("CPLUS: 點擊 LOGIN 按鈕...", flush=True)
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='LOGIN']]")))  # 優化選擇器
    ActionChains(driver).move_to_element(login_button).click().perform()
    print("CPLUS: LOGIN 按鈕點擊成功", flush=True)
    wait_for_ajax(driver)
# CPLUS Container Movement Log
def process_cplus_movement(driver, wait, initial_files):
    print("CPLUS: 直接前往 Container Movement Log...", flush=True)
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "form")))
    wait_for_ajax(driver)
    print("CPLUS: Container Movement Log 頁面加載完成", flush=True)
    print("CPLUS: 點擊 Search...", flush=True)
    local_initial = initial_files.copy()
    start_time = time.time()
    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Search']]")))
    ActionChains(driver).move_to_element(search_button).click().perform()
    print("CPLUS: Search 按鈕點擊成功", flush=True)
    wait_for_ajax(driver)
    print("CPLUS: 點擊 Download...", flush=True)
    download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Download')]")))  # 優化選擇器
    ActionChains(driver).move_to_element(download_button).click().perform()
    print("CPLUS: Download 按鈕點擊成功", flush=True)
    try:
        driver.execute_script("arguments[0].click();", download_button)
        print("CPLUS: Download 按鈕 JavaScript 點擊成功", flush=True)
    except Exception as js_e:
        print(f"CPLUS: Download 按鈕 JavaScript 點擊失敗: {str(js_e)}", flush=True)
    new_files = wait_for_new_file(local_initial, start_time, timeout=10, expected_pattern=None)
    filtered_files = {f for f in new_files if "ContainerDetailReport" not in f}
    if len(filtered_files) >= CPLUS_MOVEMENT_COUNT:
        print(f"CPLUS: Container Movement Log 下載完成，檔案位於: {download_dir}", flush=True)
        for file in filtered_files:
            print(f"CPLUS: 新下載檔案: {file}", flush=True)
        return filtered_files
    else:
        raise Exception("CPLUS: Container Movement Log 未觸發足夠新文件下載")
# CPLUS OnHandContainerList
def process_cplus_onhand(driver, wait, initial_files):
    print("CPLUS: 前往 OnHandContainerList 頁面...", flush=True)
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/OnHandContainerList")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    wait_for_ajax(driver)
    print("CPLUS: OnHandContainerList 頁面加載完成", flush=True)
    print("CPLUS: 點擊 Search...", flush=True)
    local_initial = initial_files.copy()
    start_time = time.time()
    search_button_onhand = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Search']]")))
    ActionChains(driver).move_to_element(search_button_onhand).click().perform()
    print("CPLUS: Search 按鈕點擊成功", flush=True)
    wait_for_ajax(driver)
    print("CPLUS: 點擊 Export...", flush=True)
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Export')]")))
    ActionChains(driver).move_to_element(export_button).click().perform()
    print("CPLUS: Export 按鈕點擊成功", flush=True)
    wait_for_ajax(driver)
    print("CPLUS: 點擊 Export as CSV...", flush=True)
    export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[text()='Export as CSV']")))
    ActionChains(driver).move_to_element(export_csv_button).click().perform()
    print("CPLUS: Export as CSV 按鈕點擊成功", flush=True)
    new_files = wait_for_new_file(local_initial, start_time, timeout=10, expected_pattern=None)
    filtered_files = {f for f in new_files if "ContainerDetailReport" not in f}
    if len(filtered_files) >= CPLUS_ONHAND_COUNT:
        print(f"CPLUS: OnHandContainerList 下載完成，檔案位於: {download_dir}", flush=True)
        for file in filtered_files:
            print(f"CPLUS: 新下載檔案: {file}", flush=True)
        return filtered_files
    else:
        raise Exception("CPLUS: OnHandContainerList 未觸發足夠新文件下載")
# CPLUS Housekeeping Reports
def process_cplus_house(driver, wait, initial_files):
    print("CPLUS: 前往 Housekeeping Reports 頁面...", flush=True)
    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    wait_for_ajax(driver)
    print("CPLUS: Housekeeping Reports 頁面加載完成", flush=True)
    print("CPLUS: 等待表格加載...", flush=True)
    wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
    print("CPLUS: 表格加載完成", flush=True)
    print("CPLUS: 定位並點擊所有 Excel 下載按鈕...", flush=True)
    local_initial = initial_files.copy()
    start_time = time.time()
    excel_buttons = driver.find_elements(By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and not(@disabled) and descendant::svg[@data-testid='DownloadIcon']]")  # 優化為更穩定選擇器
    button_count = len(excel_buttons)
    print(f"CPLUS: 最終找到 {button_count} 個 Excel 下載按鈕", flush=True)
    if button_count == 0:
        print("CPLUS: 未找到任何 Excel 下載按鈕", flush=True)
        return set(), 0, 0
    new_files = set()
    for idx, button in enumerate(excel_buttons):
        driver.execute_script("arguments[0].scrollIntoView(true);", button)
        ActionChains(driver).move_to_element(button).click().perform()
        temp_new = wait_for_new_file(local_initial, start_time, timeout=10, expected_pattern=None)
        filtered_files = {f for f in temp_new if "ContainerDetailReport" not in f}
        if filtered_files:
            local_initial.update(filtered_files)
            new_files.update(filtered_files)
    if len(new_files) != button_count:
        print(f"CPLUS: Housekeeping Reports 下載文件數量 ({len(new_files)}) 不等於按鈕數量 ({button_count})", flush=True)
        raise Exception(f"CPLUS: Housekeeping Reports 下載文件數量不正確，預期 {button_count} 個")
    print(f"CPLUS: Housekeeping Reports 下載完成，共 {len(new_files)} 個文件，預期 {button_count} 個", flush=True)
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
        print("CPLUS WebDriver 初始化成功", flush=True)
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
                    print(f"CPLUS: 開始處理 {section_name}，嘗試 {attempt+1}/{MAX_RETRIES}", flush=True)
                    if section_name == 'house':
                        new_files, count, button_count = section_func(driver, wait, initial_files)
                        house_file_count = count
                        house_button_count = button_count
                    else:
                        new_files = section_func(driver, wait, initial_files)
                    downloaded_files.update(new_files)
                    initial_files.update(new_files)
                    success = True
                    print(f"CPLUS: {section_name} 處理成功，新增文件: {new_files}", flush=True)
                    break
                except Exception as e:
                    print(f"CPLUS: {section_name} 嘗試 {attempt+1}/{MAX_RETRIES} 失敗: {str(e)}", flush=True)
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(5)
            if not success:
                print(f"CPLUS: {section_name} 經過 {MAX_RETRIES} 次嘗試失敗", flush=True)
        return downloaded_files, house_file_count, house_button_count
    except Exception as e:
        print(f"CPLUS 總錯誤: {str(e)}", flush=True)
        return downloaded_files, house_file_count, house_button_count
    finally:
        if driver:
            try:
                print("CPLUS: 嘗試登出...", flush=True)
                logout_menu_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Logout']]")))  # 優化選擇器
                ActionChains(driver).move_to_element(logout_menu_button).click().perform()
                print("CPLUS: 用戶菜單點擊成功", flush=True)
                logout_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//li[text()='Logout']")))
                ActionChains(driver).move_to_element(logout_option).click().perform()
                print("CPLUS: Logout 選項點擊成功", flush=True)
                close_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Close']")))
                ActionChains(driver).move_to_element(close_button).click().perform()
                print("CPLUS: Close 按鈕點擊成功", flush=True)
                driver.get("https://cplus.hit.com.hk/frontpage/#/")
                wait.until(EC.presence_of_element_located((By.XPATH, "//button[span[text()='Login']]")))
                print("CPLUS: 登出成功，回到登入頁", flush=True)
            except Exception as e:
                print(f"CPLUS: 登出失敗: {str(e)}", flush=True)
            finally:
                driver.quit()
                print("CPLUS WebDriver 關閉", flush=True)
# Barge 登入
def barge_login(driver, wait):
    print("Barge: 嘗試打開網站 https://barge.oneport.com/login...", flush=True)
    driver.get("https://barge.oneport.com/login")
    print(f"Barge: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)
    wait_for_ajax(driver)
    print("Barge: 輸入 COMPANY ID...", flush=True)
    company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Company ID']")))
    company_id_field.send_keys("CKL")
    print("Barge: COMPANY ID 輸入完成", flush=True)
    print("Barge: 輸入 USER ID...", flush=True)
    user_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='User ID']")))
    user_id_field.send_keys("barge")
    print("Barge: USER ID 輸入完成", flush=True)
    print("Barge: 輸入 PW...", flush=True)
    password_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Password']")))
    password_field.send_keys("123456")
    print("Barge: PW 輸入完成", flush=True)
    print("Barge: 點擊 LOGIN 按鈕...", flush=True)
    login_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='LOGIN']")))
    ActionChains(driver).move_to_element(login_button_barge).click().perform()
    print("Barge: LOGIN 按鈕點擊成功", flush=True)
    wait_for_ajax(driver)
# Barge 下載部分
def process_barge_download(driver, wait, initial_files):
    print("Barge: 直接前往 https://barge.oneport.com/downloadReport...", flush=True)
    driver.get("https://barge.oneport.com/downloadReport")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    wait_for_ajax(driver)
    print("Barge: downloadReport 頁面加載完成", flush=True)
    print("Barge: 選擇 Report Type...", flush=True)
    start_time = time.time()
    report_type_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-select-trigger")))
    ActionChains(driver).move_to_element(report_type_trigger).click().perform()
    print("Barge: Report Type 選擇開始", flush=True)
    container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-option[span[text()='Container Detail']]")))
    ActionChains(driver).move_to_element(container_detail_option).click().perform()
    print("Barge: Container Detail 點擊成功", flush=True)
    local_initial = initial_files.copy()
    download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Download']]")))
    ActionChains(driver).move_to_element(download_button_barge).click().perform()
    print("Barge: Download 按鈕點擊成功", flush=True)
    new_files = wait_for_new_file(local_initial, start_time, timeout=10, expected_pattern="ContainerDetailReport")
    filtered_files = {f for f in new_files if "ContainerDetailReport" in f}
    if len(filtered_files) >= BARGE_COUNT:
        print(f"Barge: Container Detail 下載完成，檔案位於: {download_dir}", flush=True)
        for file in filtered_files:
            print(f"Barge: 新下載檔案: {file}", flush=True)
        return filtered_files
    else:
        raise Exception("Barge: Container Detail 未觸發足夠新文件下載")
# Barge 操作
def process_barge():
    driver = None
    downloaded_files = set()
    try:
        driver = webdriver.Chrome(options=get_chrome_options())
        print("Barge WebDriver 初始化成功", flush=True)
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
                print(f"Barge 下載嘗試 {attempt+1}/{MAX_RETRIES} 失敗: {str(e)}", flush=True)
                if attempt < MAX_RETRIES - 1:
                    time.sleep(5)
        if not success:
            print(f"Barge 下載經過 {MAX_RETRIES} 次嘗試失敗", flush=True)
        return downloaded_files
    except Exception as e:
        print(f"Barge 總錯誤: {str(e)}", flush=True)
        return downloaded_files
    finally:
        try:
            if driver:
                print("Barge: 點擊工具欄進行登出...", flush=True)
                logout_toolbar_barge = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[span[contains(text(), 'Logout')]]")))
                driver.execute_script("arguments[0].click();", logout_toolbar_barge)
                print("Barge: 工具欄點擊成功", flush=True)
                logout_button_barge = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Logout']]")))
                driver.execute_script("arguments[0].click();", logout_button_barge)
                print("Barge: Logout 選項點擊成功", flush=True)
        except Exception as e:
            print(f"Barge: 登出失敗: {str(e)}", flush=True)
        if driver:
            driver.quit()
            print("Barge WebDriver 關閉", flush=True)
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
    cplus_thread = threading.Thread(target=lambda: update_cplus_files_and_count(process_cplus()))
    barge_thread = threading.Thread(target=lambda: barge_files.update(process_barge()))
    cplus_thread.start()
    barge_thread.start()
    cplus_thread.join()
    barge_thread.join()
    print("檢查所有下載文件...", flush=True)
    downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
    expected_file_count = CPLUS_MOVEMENT_COUNT + CPLUS_ONHAND_COUNT + house_button_count[0] + BARGE_COUNT
    print(f"預期文件數量: {expected_file_count} (Movement: {CPLUS_MOVEMENT_COUNT}, OnHand: {CPLUS_ONHAND_COUNT}, Housekeeping: {house_button_count[0]}, Barge: {BARGE_COUNT})", flush=True)
    if house_file_count[0] != house_button_count[0]:
        print(f"Housekeeping Reports 下載文件數量（{house_file_count[0]}）不等於按鈕數量（{house_button_count[0]}），放棄發送郵件", flush=True)
        print("下載文件列表：", downloaded_files, flush=True)
        return
    print(f"總共下載 {len(downloaded_files)} 個文件:", flush=True)
    for file in downloaded_files:
        print(f"找到檔案: {file}", flush=True)
    cplus_file_count = len([f for f in downloaded_files if "ContainerDetailReport" not in f])
    barge_file_count = len([f for f in downloaded_files if "ContainerDetailReport" in f])
    if cplus_file_count < (CPLUS_MOVEMENT_COUNT + CPLUS_ONHAND_COUNT + house_button_count[0]) or barge_file_count < BARGE_COUNT:
        print(f"CPLUS 文件數量（{cplus_file_count}）或 Barge 文件數量（{barge_file_count}）不足，預期 CPLUS: {CPLUS_MOVEMENT_COUNT + CPLUS_ONHAND_COUNT + house_button_count[0]}，Barge: {BARGE_COUNT}", flush=True)
        print("下載文件列表：", downloaded_files, flush=True)
        return
    if len(downloaded_files) == expected_file_count:
        print(f"所有下載完成，檔案位於: {download_dir}", flush=True)
        print("開始發送郵件...", flush=True)
        try:
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
            print("郵件發送成功!", flush=True)
        except Exception as e:
            print(f"郵件發送失敗: {str(e)}", flush=True)
    else:
        print(f"總下載文件數量不足（{len(downloaded_files)}/{expected_file_count}），放棄發送郵件", flush=True)
        print("下載文件列表：", downloaded_files, flush=True)
    print("腳本完成", flush=True)
if __name__ == "__main__":
    setup_environment()
    main()
