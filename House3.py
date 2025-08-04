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
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# 全局變量
download_dir = os.path.abspath("downloads")
CPLUS_MOVEMENT_COUNT = 1  # Container Movement Log
CPLUS_ONHAND_COUNT = 1  # OnHandContainerList
BARGE_COUNT = 1  # Barge
MAX_RETRIES = 3
DOWNLOAD_TIMEOUT = 15  # Increased download wait timeout

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

# 安全點擊元素
def safe_click(driver, element):
    methods = ['action', 'js', 'normal']
    for method in methods:
        try:
            if method == 'action':
                ActionChains(driver).move_to_element(element).click().perform()
            elif method == 'js':
                driver.execute_script("arguments[0].click();", element)
            elif method == 'normal':
                element.click()
            return True
        except Exception as e:
            print(f"點擊方法 {method} 失敗: {str(e)}", flush=True)
    return False

# 檢查新文件出現
def wait_for_new_file(initial_files, timeout=DOWNLOAD_TIMEOUT):
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_files = set(f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx')))
        new_files = current_files - initial_files
        if new_files:
            return new_files
        time.sleep(0.5)
    return set()

# CPLUS 登入
def cplus_login(driver, wait):
    print("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/", flush=True)
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    print(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)
    print("CPLUS: 點擊登錄前按鈕...", flush=True)
    login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
    safe_click(driver, login_button_pre)
    print("CPLUS: 登錄前按鈕點擊成功", flush=True)
    print("CPLUS: 輸入 COMPANY CODE...", flush=True)
    company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    company_code_field.send_keys("CKL")
    print("CPLUS: COMPANY CODE 輸入完成", flush=True)
    print("CPLUS: 輸入 USER ID...", flush=True)
    user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
    user_id_field.send_keys("KEN")
    print("CPLUS: USER ID 輸入完成", flush=True)
    print("CPLUS: 輸入 PASSWORD...", flush=True)
    password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
    password_field.send_keys(os.environ.get('SITE_PASSWORD', ''))
    print("CPLUS: PASSWORD 輸入完成", flush=True)
    print("CPLUS: 點擊 LOGIN 按鈕...", flush=True)
    login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
    safe_click(driver, login_button)
    print("CPLUS: LOGIN 按鈕點擊成功", flush=True)

# CPLUS Container Movement Log
def process_cplus_movement(driver, wait, initial_files):
    print("CPLUS: 直接前往 Container Movement Log...", flush=True)
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")))
    print("CPLUS: Container Movement Log 頁面加載完成", flush=True)
    print("CPLUS: 點擊 Search...", flush=True)
    local_initial = initial_files.copy()
    search_button = None
    for attempt in range(2):
        try:
            search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")))
            break
        except TimeoutException:
            try:
                search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and .//span[contains(text(), 'Search')]]")))
                break
            except TimeoutException:
                try:
                    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]")))
                    break
                except TimeoutException:
                    pass
    if not search_button:
        raise Exception("CPLUS: Container Movement Log Search 按鈕點擊失敗")
    safe_click(driver, search_button)
    print("CPLUS: Search 按鈕點擊成功", flush=True)
    # Wait for results to load by waiting for download button
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button")))
    print("CPLUS: 點擊 Download...", flush=True)
    download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button")))
    new_files = set()
    for attempt in range(MAX_RETRIES):
        if safe_click(driver, download_button):
            print(f"CPLUS: Download 按鈕點擊成功 (嘗試 {attempt+1})", flush=True)
            temp_new = wait_for_new_file(local_initial)
            if temp_new:
                new_files = temp_new
                break
        time.sleep(1)
    if new_files:
        filtered_files = {f for f in new_files if "cntrMoveLog" in f}
        if not filtered_files:
            raise Exception("CPLUS: Container Movement Log 未下載預期檔案")
        return filtered_files
    else:
        driver.save_screenshot("movement_download_failure.png")
        raise Exception("CPLUS: Container Movement Log 未觸發新文件下載")

# CPLUS OnHandContainerList
def process_cplus_onhand(driver, wait, initial_files):
    print("CPLUS: 前往 OnHandContainerList 頁面...", flush=True)
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/OnHandContainerList")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    print("CPLUS: OnHandContainerList 頁面加載完成", flush=True)
    print("CPLUS: 點擊 Search...", flush=True)
    local_initial = initial_files.copy()
    search_button_onhand = None
    try:
        search_button_onhand = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")))
    except TimeoutException:
        try:
            search_button_onhand = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search') or contains(@class, 'MuiButtonBase-root')]")))
        except TimeoutException:
            driver.save_screenshot("onhand_search_failure.png")
            raise Exception("CPLUS: OnHandContainerList Search 按鈕點擊失敗")
    safe_click(driver, search_button_onhand)
    print("CPLUS: Search 按鈕點擊成功", flush=True)
    # Wait for export button
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button")))
    print("CPLUS: 點擊 Export...", flush=True)
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button")))
    safe_click(driver, export_button)
    print("CPLUS: Export 按鈕點擊成功", flush=True)
    print("CPLUS: 點擊 Export as CSV...", flush=True)
    export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']")))
    new_files = set()
    for attempt in range(MAX_RETRIES):
        if safe_click(driver, export_csv_button):
            print(f"CPLUS: Export as CSV 按鈕點擊成功 (嘗試 {attempt+1})", flush=True)
            temp_new = wait_for_new_file(local_initial)
            if temp_new:
                new_files = temp_new
                break
        time.sleep(1)
    if new_files:
        # Make filter dynamic for year
        current_year = datetime.now().strftime('%Y')
        filtered_files = {f for f in new_files if f"data_{current_year}" in f}
        if not filtered_files:
            raise Exception("CPLUS: OnHandContainerList 未下載預期檔案")
        return filtered_files
    else:
        driver.save_screenshot("onhand_download_failure.png")
        raise Exception("CPLUS: OnHandContainerList 未觸發新文件下載")

# CPLUS Housekeeping Reports
def process_cplus_house(driver, wait, initial_files):
    print("CPLUS: 前往 Housekeeping Reports 頁面...", flush=True)
    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    print("CPLUS: Housekeeping Reports 頁面加載完成", flush=True)
    try:
        wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
        print("CPLUS: 表格加載完成", flush=True)
    except TimeoutException:
        driver.refresh()
        wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
        print("CPLUS: 表格加載完成 (after refresh)", flush=True)
    print("CPLUS: 定位並點擊所有 Excel 下載按鈕...", flush=True)
    local_initial = initial_files.copy()
    new_files = set()
    excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)]")
    if not excel_buttons:
        excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]//svg[@viewBox='0 0 24 24']//path[@fill='#036e11']")
    button_count = len(excel_buttons)
    print(f"CPLUS: 找到 {button_count} 個 Excel 下載按鈕", flush=True)
    downloaded_count = 0
    for idx, button in enumerate(excel_buttons):
        success = False
        report_name = "未知"
        try:
            report_name = driver.find_element(By.XPATH, f"//table[contains(@class, 'MuiTable-root')]//tbody//tr[{idx+1}]//td[3]").text
        except:
            pass
        for button_attempt in range(MAX_RETRIES):
            print(f"CPLUS: 準備點擊第 {idx+1} 個 Excel 按鈕 (嘗試 {button_attempt+1}/{MAX_RETRIES})，報告名稱: {report_name}", flush=True)
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", button)
                time.sleep(2)
                if safe_click(driver, button):
                    time.sleep(2)
                    temp_new = wait_for_new_file(local_initial, timeout=DOWNLOAD_TIMEOUT)
                    if temp_new:
                        print(f"CPLUS: 第 {idx+1} 個按鈕下載新文件: {', '.join(temp_new)}", flush=True)
                        local_initial.update(temp_new)
                        new_files.update(temp_new)
                        success = True
                        downloaded_count += 1
                        break
            except Exception as e:
                print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕嘗試失敗: {str(e)}", flush=True)
        if not success:
            print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕失敗", flush=True)
            driver.save_screenshot(f"house_button_{idx+1}_failure.png")
        time.sleep(3)  # Wait between buttons
    if new_files:
        print(f"CPLUS: Housekeeping Reports 下載完成，共 {downloaded_count} 個成功下載，預期 {button_count} 個", flush=True)
        return new_files, downloaded_count, button_count
    else:
        driver.save_screenshot("house_download_failure.png")
        raise Exception("CPLUS: Housekeeping Reports 未下載任何文件")

# CPLUS 操作
def process_cplus():
    downloaded_files = set()
    initial_files = set()
    house_file_count = 0
    house_button_count = 0
    driver = webdriver.Chrome(options=get_chrome_options())
    try:
        print("CPLUS WebDriver 初始化成功", flush=True)
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
                    print(f"CPLUS {section_name} 嘗試 {attempt+1}/{MAX_RETRIES} 失敗: {str(e)}", flush=True)
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(5)
            if not success:
                print(f"CPLUS {section_name} 經過 {MAX_RETRIES} 次嘗試失敗", flush=True)
        return downloaded_files, house_file_count, house_button_count
    finally:
        try:
            print("CPLUS: 嘗試登出...", flush=True)
            logout_menu_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
            safe_click(driver, logout_menu_button)
            print("CPLUS: 用戶菜單點擊成功", flush=True)
            logout_option = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Logout')]")))
            safe_click(driver, logout_option)
            print("CPLUS: Logout 選項點擊成功", flush=True)
        except Exception as e:
            print(f"CPLUS: 登出失敗: {str(e)}", flush=True)
        driver.quit()
        print("CPLUS WebDriver 關閉", flush=True)

# Barge 登入
def barge_login(driver, wait):
    print("Barge: 嘗試打開網站 https://barge.oneport.com/login...", flush=True)
    driver.get("https://barge.oneport.com/login")
    print(f"Barge: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)
    print("Barge: 輸入 COMPANY ID...", flush=True)
    company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Company ID' or contains(@id, 'mat-input-0')]")))
    company_id_field.send_keys("CKL")
    print("Barge: COMPANY ID 輸入完成", flush=True)
    print("Barge: 輸入 USER ID...", flush=True)
    user_id_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='User ID' or contains(@id, 'mat-input-1')]")
    user_id_field.send_keys("barge")
    print("Barge: USER ID 輸入完成", flush=True)
    print("Barge: 輸入 PW...", flush=True)
    password_field = driver.find_element(By.XPATH, "//input[contains(@id, 'mat-input') and @placeholder='Password' or contains(@id, 'mat-input-2')]")
    password_field.send_keys("123456")
    print("Barge: PW 輸入完成", flush=True)
    print("Barge: 點擊 LOGIN 按鈕...", flush=True)
    login_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'LOGIN') or contains(@class, 'mat-raised-button')]")))
    safe_click(driver, login_button_barge)
    print("Barge: LOGIN 按鈕點擊成功", flush=True)

# Barge 下載部分
def process_barge_download(driver, wait, initial_files):
    print("Barge: 直接前往 https://barge.oneport.com/downloadReport...", flush=True)
    driver.get("https://barge.oneport.com/downloadReport")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    print("Barge: downloadReport 頁面加載完成", flush=True)
    print("Barge: 選擇 Report Type...", flush=True)
    report_type_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-form-field[.//mat-label[contains(text(), 'Report Type')]]//div[contains(@class, 'mat-select-trigger')]")))
    safe_click(driver, report_type_trigger)
    print("Barge: Report Type 選擇開始", flush=True)
    print("Barge: 點擊 Container Detail...", flush=True)
    container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-option//span[contains(text(), 'Container Detail')]")))
    safe_click(driver, container_detail_option)
    print("Barge: Container Detail 點擊成功", flush=True)
    # Wait for download button
    wait.until(EC.presence_of_element_located((By.XPATH, "//button[span[text()='Download']]")))
    print("Barge: 點擊 Download...", flush=True)
    local_initial = initial_files.copy()
    download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Download']]")))
    new_files = set()
    for attempt in range(MAX_RETRIES):
        if safe_click(driver, download_button_barge):
            print(f"Barge: Download 按鈕點擊成功 (嘗試 {attempt+1})", flush=True)
            temp_new = wait_for_new_file(local_initial)
            if temp_new:
                new_files = temp_new
                break
        time.sleep(1)
    if new_files:
        filtered_files = {f for f in new_files if "ContainerDetailReport" in f}
        if not filtered_files:
            raise Exception("Barge: Container Detail 未下載預期檔案")
        return filtered_files
    else:
        driver.save_screenshot("barge_download_failure.png")
        raise Exception("Barge: Container Detail 未觸發新文件下載")

# Barge 操作
def process_barge():
    downloaded_files = set()
    driver = webdriver.Chrome(options=get_chrome_options())
    try:
        print("Barge WebDriver 初始化成功", flush=True)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        wait = WebDriverWait(driver, 20)
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
    finally:
        try:
            print("Barge: 點擊工具欄進行登出...", flush=True)
            logout_toolbar_barge = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
            driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar_barge)
            safe_click(driver, logout_toolbar_barge)
            print("Barge: 工具欄點擊成功", flush=True)
            logout_span_xpath = "//div[contains(@class, 'mat-menu-panel')]//button//span[contains(text(), 'Logout')]"
            logout_button_barge = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, logout_span_xpath)))
            driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
            safe_click(driver, logout_button_barge)
            print("Barge: Logout 選項點擊成功", flush=True)
        except Exception as e:
            print(f"Barge: 登出失敗: {str(e)}", flush=True)
        driver.quit()
        print("Barge WebDriver 關閉", flush=True)

# 主函數
def main():
    clear_download_dir()
    cplus_files, house_file_count, house_button_count = process_cplus()
    barge_files = process_barge()
    print("檢查所有下載文件...", flush=True)
    downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
    expected_file_count = CPLUS_MOVEMENT_COUNT + CPLUS_ONHAND_COUNT + house_button_count + BARGE_COUNT
    print(f"預期文件數量: {expected_file_count} (Movement: {CPLUS_MOVEMENT_COUNT}, OnHand: {CPLUS_ONHAND_COUNT}, Housekeeping: {house_button_count}, Barge: {BARGE_COUNT})", flush=True)
    if house_file_count < house_button_count:
        print(f"Housekeeping Reports 下載文件數量（{house_file_count}）少於按鈕數量（{house_button_count}），放棄發送郵件", flush=True)
        return
    if len(downloaded_files) >= expected_file_count:
        print(f"所有下載完成，檔案位於: {download_dir}", flush=True)
        for file in downloaded_files:
            print(f"找到檔案: {file}", flush=True)
        print("開始發送郵件...", flush=True)
        try:
            smtp_server = 'smtp.zoho.com'
            smtp_port = 587
            sender_email = os.environ.get('ZOHO_EMAIL', 'paklun_ckline@zohomail.com')
            sender_password = os.environ.get('ZOHO_PASSWORD', '@d6G.Pie5UkEPqm')
            receiver_email = 'ckeqc@ckline.com.hk'
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
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, receiver_email, msg.as_string())
            print("郵件發送成功!", flush=True)
        except Exception as e:
            print(f"郵件發送失敗: {str(e)}", flush=True)
    else:
        print(f"總下載文件數量不足（{len(downloaded_files)}/{expected_file_count}），放棄發送郵件", flush=True)
    print("腳本完成", flush=True)

if __name__ == "__main__":
    setup_environment()
    main()
