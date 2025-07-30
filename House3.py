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
EXPECTED_FILE_COUNT = 9  # 預期文件數：1 (Container Movement Log) + 1 (OnHandContainerList) + 6 (Housekeeping Reports) + 1 (Barge)
EXPECTED_HOUSEKEEPING_FILES = {
    "CONTAINER DAMAGE REPORT (LINE) ENTRY GATE + EXIT GATE": "DM1C",
    "CY - GATELOG": "GA1",
    "CONTAINER LIST (ON HAND)": "IA15",
    "CONTAINER LIST (DAMAGED)": "IA17",
    "ACTIVE REEFER CONTAINER ON HAND LIST": "IA5",
    "REEFER CONTAINER MONITOR REPORT": "IE2"
}

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

# 檢查新文件出現
def wait_for_new_file(initial_files, timeout=20):
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_files = set(f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx')))
        new_files = current_files - initial_files
        if new_files:
            return new_files
        time.sleep(0.1)
    return set()

# CPLUS 登錄
def cplus_login(driver):
    try:
        wait = WebDriverWait(driver, 20)
        print("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/", flush=True)
        driver.get("https://cplus.hit.com.hk/frontpage/#/")
        print(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)
        time.sleep(2)

        print("CPLUS: 點擊登錄前按鈕...", flush=True)
        login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
        ActionChains(driver).move_to_element(login_button_pre).click().perform()
        print("CPLUS: 登錄前按鈕點擊成功", flush=True)
        time.sleep(2)

        print("CPLUS: 輸入 COMPANY CODE...", flush=True)
        company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
        company_code_field.send_keys("CKL")
        print("CPLUS: COMPANY CODE 輸入完成", flush=True)
        time.sleep(1)

        print("CPLUS: 輸入 USER ID...", flush=True)
        user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
        user_id_field.send_keys("KEN")
        print("CPLUS: USER ID 輸入完成", flush=True)
        time.sleep(1)

        print("CPLUS: 輸入 PASSWORD...", flush=True)
        password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
        password_field.send_keys(os.environ.get('SITE_PASSWORD'))
        print("CPLUS: PASSWORD 輸入完成", flush=True)
        time.sleep(1)

        print("CPLUS: 點擊 LOGIN 按鈕...", flush=True)
        login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
        ActionChains(driver).move_to_element(login_button).click().perform()
        print("CPLUS: LOGIN 按鈕點擊成功", flush=True)
        time.sleep(2)
        return True
    except Exception as e:
        print(f"CPLUS 登錄失敗: {str(e)}", flush=True)
        return False

# CPLUS 登出
def cplus_logout(driver):
    try:
        wait = WebDriverWait(driver, 30)
        print("CPLUS: 嘗試登出...", flush=True)
        logout_menu_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", logout_menu_button)
        time.sleep(0.5)
        ActionChains(driver).move_to_element(logout_menu_button).click().perform()
        print("CPLUS: 登錄按鈕點擊成功", flush=True)

        logout_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='menu-list-grow']/div[6]/li")))
        driver.execute_script("arguments[0].scrollIntoView(true);", logout_option)
        time.sleep(0.5)
        ActionChains(driver).move_to_element(logout_option).click().perform()
        print("CPLUS: Logout 選項點擊成功", flush=True)
        time.sleep(2)
    except TimeoutException:
        print("CPLUS: 登出按鈕未找到，嘗試備用定位...", flush=True)
        try:
            logout_option = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Logout')]")))
            ActionChains(driver).move_to_element(logout_option).click().perform()
            print("CPLUS: 備用 Logout 選項點擊成功", flush=True)
            time.sleep(2)
        except TimeoutException:
            print("CPLUS: 備用 Logout 選項未找到，跳過登出", flush=True)
    except Exception as e:
        print(f"CPLUS: 登出失敗: {str(e)}", flush=True)

# 下載 Container Movement Log
def download_container_movement_log(driver, initial_files):
    downloaded_files = set()
    try:
        print("CPLUS: 直接前往 Container Movement Log...", flush=True)
        driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
        time.sleep(2)
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")))
        print("CPLUS: Container Movement Log 頁面加載完成", flush=True)

        print("CPLUS: 點擊 Search...", flush=True)
        for attempt in range(2):
            try:
                search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")))
                ActionChains(driver).move_to_element(search_button).click().perform()
                print("CPLUS: Search 按鈕點擊成功", flush=True)
                break
            except TimeoutException:
                print(f"CPLUS: Search 按鈕未找到，嘗試備用定位 {attempt+1}/2...", flush=True)
                try:
                    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and .//span[contains(text(), 'Search')]]")))
                    ActionChains(driver).move_to_element(search_button).click().perform()
                    print("CPLUS: 備用 Search 按鈕 1 點擊成功", flush=True)
                    break
                except TimeoutException:
                    print(f"CPLUS: 備用 Search 按鈕 1 失敗，嘗試備用定位 2 (嘗試 {attempt+1}/2)...", flush=True)
                    try:
                        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]")))
                        ActionChains(driver).move_to_element(search_button).click().perform()
                        print("CPLUS: 備用 Search 按鈕 2 點擊成功", flush=True)
                        break
                    except TimeoutException:
                        print(f"CPLUS: 備用 Search 按鈕 2 失敗 (嘗試 {attempt+1}/2)", flush=True)
        else:
            print("CPLUS: Container Movement Log Search 按鈕點擊失敗，重試 2 次後放棄", flush=True)
            return downloaded_files, False

        print("CPLUS: 點擊 Download...", flush=True)
        for attempt in range(2):
            try:
                download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button")))
                ActionChains(driver).move_to_element(download_button).click().perform()
                print("CPLUS: Download 按鈕點擊成功", flush=True)
                time.sleep(0.5)
                try:
                    driver.execute_script("arguments[0].click();", download_button)
                    print("CPLUS: Download 按鈕 JavaScript 點擊成功", flush=True)
                except Exception as js_e:
                    print(f"CPLUS: Download 按鈕 JavaScript 點擊失敗: {str(js_e)}", flush=True)
                time.sleep(0.5)
                break
            except Exception as e:
                print(f"CPLUS: Download 按鈕點擊失敗 (嘗試 {attempt+1}/2): {str(e)}", flush=True)
                time.sleep(0.5)
        else:
            print("CPLUS: Container Movement Log Download 按鈕點擊失敗，重試 2 次後放棄", flush=True)
            return downloaded_files, False

        new_files = wait_for_new_file(initial_files, timeout=20)
        if new_files and any("cntrMoveLog.xlsx" in f for f in new_files):
            print(f"CPLUS: Container Movement Log 下載完成，檔案位於: {download_dir}", flush=True)
            for file in new_files:
                print(f"CPLUS: 新下載檔案: {file}", flush=True)
            downloaded_files.update(new_files)
            return downloaded_files, True
        else:
            print("CPLUS: Container Movement Log 未觸發新文件下載", flush=True)
            return downloaded_files, False

# 下載 OnHandContainerList
def download_onhand_container_list(driver, initial_files):
    downloaded_files = set()
    try:
        print("CPLUS: 前往 OnHandContainerList 頁面...", flush=True)
        driver.get("https://cplus.hit.com.hk/app/#/enquiry/OnHandContainerList")
        time.sleep(2)
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
        print("CPLUS: OnHandContainerList 頁面加載完成", flush=True)

        print("CPLUS: 點擊 Search...", flush=True)
        for attempt in range(2):
            try:
                search_button_onhand = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")))
                ActionChains(driver).move_to_element(search_button_onhand).click().perform()
                print("CPLUS: Search 按鈕點擊成功", flush=True)
                break
            except TimeoutException:
                print(f"CPLUS: Search 按鈕未找到，嘗試備用定位 {attempt+1}/2...", flush=True)
                try:
                    search_button_onhand = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]")))
                    ActionChains(driver).move_to_element(search_button_onhand).click().perform()
                    print("CPLUS: 備用 Search 按鈕點擊成功", flush=True)
                    break
                except TimeoutException:
                    print(f"CPLUS: 備用 Search 按鈕失敗 (嘗試 {attempt+1}/2)", flush=True)
        else:
            print("CPLUS: OnHandContainerList Search 按鈕點擊失敗，重試 2 次後放棄", flush=True)
            return downloaded_files, False

        print("CPLUS: 點擊 Export...", flush=True)
        for attempt in range(2):
            try:
                export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button")))
                ActionChains(driver).move_to_element(export_button).click().perform()
                print("CPLUS: Export 按鈕點擊成功", flush=True)
                time.sleep(0.5)
                break
            except Exception as e:
                print(f"CPLUS: Export 按鈕點擊失敗 (嘗試 {attempt+1}/2): {str(e)}", flush=True)
                time.sleep(0.5)
        else:
            print("CPLUS: OnHandContainerList Export 按鈕點擊失敗，重試 2 次後放棄", flush=True)
            return downloaded_files, False

        print("CPLUS: 點擊 Export as CSV...", flush=True)
        for attempt in range(2):
            try:
                export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']")))
                ActionChains(driver).move_to_element(export_csv_button).click().perform()
                print("CPLUS: Export as CSV 按鈕點擊成功", flush=True)
                time.sleep(0.5)
                break
            except Exception as e:
                print(f"CPLUS: Export as CSV 按鈕點擊失敗 (嘗試 {attempt+1}/2): {str(e)}", flush=True)
                time.sleep(0.5)
        else:
            print("CPLUS: OnHandContainerList Export as CSV 按鈕點擊失敗，重試 2 次後放棄", flush=True)
            return downloaded_files, False

        new_files = wait_for_new_file(initial_files, timeout=20)
        if new_files and any("data_" in f for f in new_files):
            print(f"CPLUS: OnHandContainerList 下載完成，檔案位於: {download_dir}", flush=True)
            for file in new_files:
                print(f"CPLUS: 新下載檔案: {file}", flush=True)
            downloaded_files.update(new_files)
            return downloaded_files, True
        else:
            print("CPLUS: OnHandContainerList 未觸發新文件下載", flush=True)
            return downloaded_files, False

# 下載 Housekeeping Reports
def download_housekeeping_reports(driver, initial_files):
    downloaded_files = set()
    failed_buttons = []
    try:
        print("CPLUS: 前往 Housekeeping Reports 頁面...", flush=True)
        driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
        time.sleep(2)
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
        print("CPLUS: Housekeeping Reports 頁面加載完成", flush=True)

        print("CPLUS: 等待表格加載...", flush=True)
        try:
            wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
            print("CPLUS: 表格加載完成", flush=True)
        except TimeoutException:
            print("CPLUS: 表格未加載，跳過下載步驟", flush=True)
            return downloaded_files, False

        print("CPLUS: 定位並點擊所有 Excel 下載按鈕...", flush=True)
        excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]//svg[@viewBox='0 0 24 24']//path[@fill='#036e11']")
        button_count = len(excel_buttons)
        print(f"CPLUS: 找到 {button_count} 個 Excel 下載按鈕", flush=True)

        if button_count == 0:
            print("CPLUS: 未找到 Excel 按鈕，嘗試備用定位...", flush=True)
            excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)]")
            button_count = len(excel_buttons)
            print(f"CPLUS: 備用定位找到 {button_count} 個 Excel 下載按鈕", flush=True)

        for idx, button in enumerate(excel_buttons):
            try:
                button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"(//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)])[{idx+1}]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", button)
                time.sleep(0.5)

                try:
                    report_name = driver.find_element(By.XPATH, f"//table[contains(@class, 'MuiTable-root')]//tbody//tr[{idx+1}]//td[3]").text
                    expected_prefix = EXPECTED_HOUSEKEEPING_FILES.get(report_name, "Unknown")
                    print(f"CPLUS: 準備點擊第 {idx+1} 個 Excel 按鈕，報告名稱: {report_name}，預期文件前綴: {expected_prefix}", flush=True)
                except:
                    report_name = "Unknown"
                    expected_prefix = "Unknown"
                    print(f"CPLUS: 無法獲取第 {idx+1} 個按鈕的報告名稱", flush=True)

                for attempt in range(2):
                    try:
                        ActionChains(driver).move_to_element(button).pause(0.5).click().perform()
                        print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕點擊成功", flush=True)
                        break
                    except Exception as e:
                        print(f"CPLUS: 第 {idx+1} 個按鈕點擊嘗試 {attempt+1} 失敗: {str(e)}", flush=True)
                        time.sleep(0.5)
                else:
                    print(f"CPLUS: 第 {idx+1} 個按鈕點擊失敗，重試 2 次後放棄", flush=True)
                    failed_buttons.append((idx, report_name, expected_prefix))
                    continue

                new_files = wait_for_new_file(initial_files, timeout=20)
                if new_files and any(expected_prefix in f for f in new_files if expected_prefix != "Unknown"):
                    print(f"CPLUS: 第 {idx+1} 個按鈕下載新文件: {', '.join(new_files)}", flush=True)
                    initial_files.update(new_files)
                    downloaded_files.update(new_files)
                else:
                    print(f"CPLUS: 第 {idx+1} 個按鈕未觸發新文件下載", flush=True)
                    failed_buttons.append((idx, report_name, expected_prefix))

                try:
                    driver.execute_script("arguments[0].click();", button)
                    print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕 JavaScript 點擊成功", flush=True)
                    new_files = wait_for_new_file(initial_files, timeout=20)
                    if new_files and any(expected_prefix in f for f in new_files if expected_prefix != "Unknown"):
                        print(f"CPLUS: 第 {idx+1} 個按鈕下載新文件 (JavaScript): {', '.join(new_files)}", flush=True)
                        initial_files.update(new_files)
                        downloaded_files.update(new_files)
                    else:
                        print(f"CPLUS: 第 {idx+1} 個按鈕 JavaScript 未觸發新文件下載", flush=True)
                        if (idx, report_name, expected_prefix) not in failed_buttons:
                            failed_buttons.append((idx, report_name, expected_prefix))
                except Exception as js_e:
                    print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕 JavaScript 點擊失敗: {str(js_e)}", flush=True)
                    if (idx, report_name, expected_prefix) not in failed_buttons:
                        failed_buttons.append((idx, report_name, expected_prefix))

            except Exception as e:
                print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕處理失敗: {str(e)}", flush=True)
                failed_buttons.append((idx, report_name, expected_prefix))

        success = len(downloaded_files) >= 6 and all(any(prefix in f for f in downloaded_files) for prefix in EXPECTED_HOUSEKEEPING_FILES.values())
        return downloaded_files, success, failed_buttons

# 重試 Housekeeping Reports 的失敗按鈕
def retry_housekeeping_reports(driver, initial_files, failed_buttons):
    downloaded_files = set()
    still_failed = []
    try:
        print("CPLUS: 重試 Housekeeping Reports 失敗的按鈕...", flush=True)
        for idx, report_name, expected_prefix in failed_buttons:
            try:
                button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"(//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)])[{idx+1}]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", button)
                time.sleep(0.5)

                print(f"CPLUS: 重試第 {idx+1} 個 Excel 按鈕，報告名稱: {report_name}，預期文件前綴: {expected_prefix}", flush=True)
                for attempt in range(2):
                    try:
                        ActionChains(driver).move_to_element(button).pause(0.5).click().perform()
                        print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕重試點擊成功", flush=True)
                        break
                    except Exception as e:
                        print(f"CPLUS: 第 {idx+1} 個按鈕重試點擊嘗試 {attempt+1} 失敗: {str(e)}", flush=True)
                        time.sleep(0.5)
                else:
                    print(f"CPLUS: 第 {idx+1} 個按鈕重試點擊失敗，重試 2 次後放棄", flush=True)
                    still_failed.append((idx, report_name, expected_prefix))
                    continue

                new_files = wait_for_new_file(initial_files, timeout=20)
                if new_files and any(expected_prefix in f for f in new_files if expected_prefix != "Unknown"):
                    print(f"CPLUS: 第 {idx+1} 個按鈕重試下載新文件: {', '.join(new_files)}", flush=True)
                    initial_files.update(new_files)
                    downloaded_files.update(new_files)
                else:
                    print(f"CPLUS: 第 {idx+1} 個按鈕重試未觸發新文件下載", flush=True)
                    still_failed.append((idx, report_name, expected_prefix))

                try:
                    driver.execute_script("arguments[0].click();", button)
                    print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕重試 JavaScript 點擊成功", flush=True)
                    new_files = wait_for_new_file(initial_files, timeout=20)
                    if new_files and any(expected_prefix in f for f in new_files if expected_prefix != "Unknown"):
                        print(f"CPLUS: 第 {idx+1} 個按鈕重試下載新文件 (JavaScript): {', '.join(new_files)}", flush=True)
                        initial_files.update(new_files)
                        downloaded_files.update(new_files)
                    else:
                        print(f"CPLUS: 第 {idx+1} 個按鈕重試 JavaScript 未觸發新文件下載", flush=True)
                        if (idx, report_name, expected_prefix) not in still_failed:
                            still_failed.append((idx, report_name, expected_prefix))
                except Exception as js_e:
                    print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕重試 JavaScript 點擊失敗: {str(js_e)}", flush=True)
                    if (idx, report_name, expected_prefix) not in still_failed:
                        still_failed.append((idx, report_name, expected_prefix))

            except Exception as e:
                print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕重試處理失敗: {str(e)}", flush=True)
                still_failed.append((idx, report_name, expected_prefix))

        success = len(downloaded_files) >= 6 and all(any(prefix in f for f in downloaded_files) for prefix in EXPECTED_HOUSEKEEPING_FILES.values())
        return downloaded_files, success, still_failed

# CPLUS 主流程
def process_cplus():
    driver = None
    all_downloaded_files = set()
    max_retries = 3
    try:
        driver = webdriver.Chrome(options=get_chrome_options())
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        if not cplus_login(driver):
            return all_downloaded_files

        # 下載 Container Movement Log
        cm_success = False
        for attempt in range(max_retries):
            print(f"CPLUS: Container Movement Log 第 {attempt+1}/{max_retries} 次嘗試...", flush=True)
            cm_files, cm_success = download_container_movement_log(driver, all_downloaded_files)
            all_downloaded_files.update(cm_files)
            if cm_success:
                break
            if attempt < max_retries - 1:
                print(f"CPLUS: Container Movement Log 失敗，等待 5 秒後重試...", flush=True)
                time.sleep(5)

        # 下載 OnHandContainerList
        oh_success = False
        for attempt in range(max_retries):
            print(f"CPLUS: OnHandContainerList 第 {attempt+1}/{max_retries} 次嘗試...", flush=True)
            oh_files, oh_success = download_onhand_container_list(driver, all_downloaded_files)
            all_downloaded_files.update(oh_files)
            if oh_success:
                break
            if attempt < max_retries - 1:
                print(f"CPLUS: OnHandContainerList 失敗，等待 5 秒後重試...", flush=True)
                time.sleep(5)

        # 下載 Housekeeping Reports
        hr_success = False
        failed_buttons = []
        for attempt in range(max_retries):
            print(f"CPLUS: Housekeeping Reports 第 {attempt+1}/{max_retries} 次嘗試...", flush=True)
            hr_files, hr_success, failed_buttons = download_housekeeping_reports(driver, all_downloaded_files)
            all_downloaded_files.update(hr_files)
            if hr_success:
                break
            if failed_buttons and attempt < max_retries - 1:
                print(f"CPLUS: Housekeeping Reports 部分按鈕失敗，準備重試 {len(failed_buttons)} 個按鈕...", flush=True)
                hr_files, hr_success, failed_buttons = retry_housekeeping_reports(driver, all_downloaded_files, failed_buttons)
                all_downloaded_files.update(hr_files)
                if hr_success:
                    break
            if attempt < max_retries - 1:
                print(f"CPLUS: Housekeeping Reports 失敗，等待 5 秒後重試...", flush=True)
                time.sleep(5)

        return all_downloaded_files

    finally:
        if driver:
            cplus_logout(driver)
            driver.quit()
            print("CPLUS WebDriver 關閉", flush=True)

# Barge 操作
def process_barge():
    driver = None
    downloaded_files = set()
    max_retries = 3
    for attempt in range(max_retries):
        print(f"Barge: 第 {attempt+1}/{max_retries} 次嘗試...", flush=True)
        try:
            driver = webdriver.Chrome(options=get_chrome_options())
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

            print("Barge: 嘗試打開網站 https://barge.oneport.com/bargeBooking...", flush=True)
            driver.get("https://barge.oneport.com/bargeBooking")
            print(f"Barge: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)
            time.sleep(3)

            print("Barge: 輸入 COMPANY ID...", flush=True)
            wait = WebDriverWait(driver, 20)
            company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='mat-input-0']")))
            company_id_field.send_keys("CKL")
            print("Barge: COMPANY ID 輸入完成", flush=True)
            time.sleep(1)

            print("Barge: 輸入 USER ID...", flush=True)
            user_id_field_barge = driver.find_element(By.XPATH, "//*[@id='mat-input-1']")
            user_id_field_barge.send_keys("barge")
            print("Barge: USER ID 輸入完成", flush=True)
            time.sleep(1)

            print("Barge: 輸入 PW...", flush=True)
            password_field_barge = driver.find_element(By.XPATH, "//*[@id='mat-input-2']")
            password_field_barge.send_keys("123456")
            print("Barge: PW 輸入完成", flush=True)
            time.sleep(1)

            print("Barge: 點擊 LOGIN 按鈕...", flush=True)
            login_button_barge = driver.find_element(By.XPATH, "//*[@id='login-form-container']/app-login-form/form/div/button")
            ActionChains(driver).move_to_element(login_button_barge).click().perform()
            print("Barge: LOGIN 按鈕點擊成功", flush=True)
            time.sleep(3)

            print("Barge: 點擊主工具欄...", flush=True)
            toolbar_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[1]/span[1]/mat-icon")))
            ActionChains(driver).move_to_element(toolbar_button).click().perform()
            print("Barge: 主工具欄點擊成功", flush=True)
            time.sleep(2)

            print("Barge: 點擊 Report...", flush=True)
            report_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-menu-panel-4']/div/button[4]/span")))
            ActionChains(driver).move_to_element(report_button).click().perform()
            print("Barge: Report 點擊成功", flush=True)
            time.sleep(2)

            print("Barge: 選擇 Report Type...", flush=True)
            report_type_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-select-value-61']/span")))
            ActionChains(driver).move_to_element(report_type_select).click().perform()
            print("Barge: Report Type 選擇開始", flush=True)
            time.sleep(2)

            print("Barge: 點擊 Container Detail...", flush=True)
            container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-option-508']/span")))
            ActionChains(driver).move_to_element(container_detail_option).click().perform()
            print("Barge: Container Detail 點擊成功", flush=True)
            time.sleep(2)

            print("Barge: 點擊 Download...", flush=True)
            initial_files = downloaded_files.copy()
            for attempt in range(2):
                try:
                    download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='content-mount']/app-download-report/div[2]/div/form/div[2]/button")))
                    ActionChains(driver).move_to_element(download_button_barge).click().perform()
                    print("Barge: Download 按鈕點擊成功", flush=True)
                    time.sleep(0.5)
                    break
                except Exception as e:
                    print(f"Barge: Download 按鈕點擊失敗 (嘗試 {attempt+1}/2): {str(e)}", flush=True)
                    time.sleep(0.5)
            else:
                print("Barge: Container Detail Download 按鈕點擊失敗，重試 2 次後放棄", flush=True)
                return downloaded_files, False

            new_files = wait_for_new_file(initial_files, timeout=20)
            if new_files and any("ContainerDetailReport" in f for f in new_files):
                print(f"Barge: Container Detail 下載完成，檔案位於: {download_dir}", flush=True)
                for file in new_files:
                    print(f"Barge: 新下載檔案: {file}", flush=True)
                downloaded_files.update(new_files)
                return downloaded_files, True
            else:
                print("Barge: Container Detail 未觸發新文件下載", flush=True)
                return downloaded_files, False

    finally:
        if driver:
            try:
                print("Barge: 點擊工具欄進行登出...", flush=True)
                logout_toolbar_barge = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar_barge)
                time.sleep(0.5)
                ActionChains(driver).move_to_element(logout_toolbar_barge).click().perform()
                print("Barge: 工具欄點擊成功", flush=True)

                print("Barge: 點擊 Logout 選項...", flush=True)
                logout_button_barge = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-menu-panel-11']/div/button/span")))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                time.sleep(0.5)
                ActionChains(driver).move_to_element(logout_button_barge).click().perform()
                print("Barge: Logout 選項點擊成功", flush=True)
                time.sleep(2)
            except TimeoutException:
                print("Barge: 登出按鈕未找到，嘗試備用定位...", flush=True)
                try:
                    logout_button_barge = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Logout')]")))
                    ActionChains(driver).move_to_element(logout_button_barge).click().perform()
                    print("Barge: 備用 Logout 選項點擊成功", flush=True)
                    time.sleep(2)
                except TimeoutException:
                    print("Barge: 備用 Logout 選項未找到，跳過登出", flush=True)
            except Exception as e:
                print(f"Barge: 登出失敗: {str(e)}", flush=True)
            driver.quit()
            print("Barge WebDriver 關閉", flush=True)

# 主函數
def main():
    max_retries = 3
    all_downloaded_files = set()
    clear_download_dir()

    # 第一次嘗試
    print("開始運行腳本，第 1 次嘗試...", flush=True)
    cplus_success = False
    barge_success = False
    cplus_files = set()
    barge_files = set()

    cplus_thread = threading.Thread(target=lambda: cplus_files.update(process_cplus()))
    barge_thread = threading.Thread(target=lambda: barge_files.update(*process_barge()))
    cplus_thread.start()
    barge_thread.start()
    cplus_thread.join()
    barge_thread.join()
    all_downloaded_files.update(cplus_files)
    all_downloaded_files.update(barge_files)

    # 檢查每個部分的成功狀態
    cm_success = any("cntrMoveLog.xlsx" in f for f in all_downloaded_files)
    oh_success = any("data_" in f for f in all_downloaded_files)
    hr_success = len([f for f in all_downloaded_files if any(prefix in f for prefix in EXPECTED_HOUSEKEEPING_FILES.values())]) >= 6
    barge_success = any("ContainerDetailReport" in f for f in all_downloaded_files)

    # 重試失敗部分
    for attempt in range(1, max_retries):
        if cm_success and oh_success and hr_success and barge_success:
            break
        print(f"\n部分下載失敗，開始第 {attempt+1}/{max_retries} 次重試...", flush=True)
        clear_download_dir()
        new_files = set()

        driver = None
        if not (cm_success and oh_success and hr_success):
            driver = webdriver.Chrome(options=get_chrome_options())
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            if not cplus_login(driver):
                print("CPLUS: 登錄失敗，跳過重試", flush=True)
                continue

        if not cm_success:
            print(f"CPLUS: 重試 Container Movement Log，第 {attempt+1}/{max_retries} 次...", flush=True)
            cm_files, cm_success = download_container_movement_log(driver, all_downloaded_files)
            new_files.update(cm_files)

        if not oh_success:
            print(f"CPLUS: 重試 OnHandContainerList，第 {attempt+1}/{max_retries} 次...", flush=True)
            oh_files, oh_success = download_onhand_container_list(driver, all_downloaded_files)
            new_files.update(oh_files)

        if not hr_success:
            print(f"CPLUS: 重試 Housekeeping Reports，第 {attempt+1}/{max_retries} 次...", flush=True)
            hr_files, hr_success, failed_buttons = download_housekeeping_reports(driver, all_downloaded_files)
            new_files.update(hr_files)
            if failed_buttons:
                print(f"CPLUS: Housekeeping Reports 仍部分失敗，重試 {len(failed_buttons)} 個按鈕...", flush=True)
                hr_files, hr_success, failed_buttons = retry_housekeeping_reports(driver, all_downloaded_files, failed_buttons)
                new_files.update(hr_files)

        if driver:
            cplus_logout(driver)
            driver.quit()

        if not barge_success:
            print(f"Barge: 重試，第 {attempt+1}/{max_retries} 次...", flush=True)
            barge_files, barge_success = process_barge()
            new_files.update(barge_files)

        all_downloaded_files.update(new_files)
        cm_success = any("cntrMoveLog.xlsx" in f for f in all_downloaded_files)
        oh_success = any("data_" in f for f in all_downloaded_files)
        hr_success = len([f for f in all_downloaded_files if any(prefix in f for prefix in EXPECTED_HOUSEKEEPING_FILES.values())]) >= 6
        barge_success = any("ContainerDetailReport" in f for f in all_downloaded_files)

        if attempt < max_retries - 1:
            print(f"等待 5 秒後進行下一次重試...", flush=True)
            time.sleep(5)

    # 最終檢查
    downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
    if len(downloaded_files) >= EXPECTED_FILE_COUNT:
        print(f"所有下載完成，檔案位於: {download_dir}", flush=True)
        for file in downloaded_files:
            print(f"找到檔案: {file}", flush=True)

        # 發送 Zoho Mail
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

            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
            server.quit()
            print("郵件發送成功!", flush=True)
        except Exception as e:
            print(f"郵件發送失敗: {str(e)}", flush=True)
    else:
        print(f"下載文件數量不足（{len(downloaded_files)}/{EXPECTED_FILE_COUNT}），放棄發送郵件", flush=True)
        for file in downloaded_files:
            print(f"找到檔案: {file}", flush=True)
        if not cm_success:
            print("失敗部分: Container Movement Log", flush=True)
        if not oh_success:
            print("失敗部分: OnHandContainerList", flush=True)
        if not hr_success:
            print("失敗部分: Housekeeping Reports", flush=True)
            for idx, report_name, expected_prefix in failed_buttons:
                print(f"  未下載報告: {report_name} (預期前綴: {expected_prefix})", flush=True)
        if not barge_success:
            print("失敗部分: Barge", flush=True)

    print("腳本完成", flush=True)

if __name__ == "__main__":
    setup_environment()
    main()
