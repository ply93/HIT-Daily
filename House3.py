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

# 全局變量
download_dir = os.path.abspath("downloads")
EXPECTED_TOTAL_FILES = 9  # 1 (Container Movement Log) + 1 (OnHandContainerList) + 6 (Housekeeping Reports) + 1 (Barge)

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
        if "selenium" not in result.stdout:
            subprocess.run(['pip', 'install', 'selenium'], check=True)
            print("Selenium 已安裝", flush=True)
        else:
            print("Selenium 已存在，跳過安裝", flush=True)
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
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": False
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.binary_location = '/usr/bin/chromium-browser'
    return chrome_options

# 檢查新文件出現，並確保沒有正在下載的文件
def wait_for_new_file(initial_files, timeout=15):
    start_time = time.time()
    while time.time() - start_time < timeout:
        all_files = set(os.listdir(download_dir))
        downloading = any(f.endswith('.crdownload') or f.endswith('.tmp') for f in all_files)
        current_files = set(f for f in all_files if f.endswith(('.csv', '.xlsx')))
        new_files = current_files - initial_files
        if new_files and not downloading:
            return new_files
        time.sleep(0.5)
    return set()

# 保存頁面源碼
def save_page_source(driver, filename):
    with open(os.path.join(download_dir, filename), 'w', encoding='utf-8') as f:
        f.write(driver.page_source)
    print(f"已保存頁面源碼: {filename}", flush=True)

# CPLUS 登錄
def cplus_login(driver, wait):
    try:
        print("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/", flush=True)
        driver.get("https://cplus.hit.com.hk/frontpage/#/")
        print(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)
        time.sleep(2)

        if "404.html" in driver.current_url:
            print("CPLUS: 檢測到 404 錯誤，保存截圖和源碼", flush=True)
            driver.save_screenshot("cplus_login_404.png")
            save_page_source(driver, "cplus_login_404.html")
            return False

        print("CPLUS: 點擊登錄前按鈕...", flush=True)
        login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
        login_button_pre.click()
        print("CPLUS: 登錄前按鈕點擊成功", flush=True)
        time.sleep(2)

        print("CPLUS: 輸入 COMPANY CODE...", flush=True)
        company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
        company_code_field.clear()
        company_code_field.send_keys("CKL")
        print("CPLUS: COMPANY CODE 輸入完成", flush=True)
        time.sleep(1)

        print("CPLUS: 輸入 USER ID...", flush=True)
        user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
        user_id_field.clear()
        user_id_field.send_keys("KEN")
        print("CPLUS: USER ID 輸入完成", flush=True)
        time.sleep(1)

        print("CPLUS: 輸入 PASSWORD...", flush=True)
        password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
        password_field.clear()
        password_field.send_keys(os.environ.get('SITE_PASSWORD'))
        print("CPLUS: PASSWORD 輸入完成", flush=True)
        time.sleep(1)

        print("CPLUS: 點擊 LOGIN 按鈕...", flush=True)
        login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
        login_button.click()
        print("CPLUS: LOGIN 按鈕點擊成功", flush=True)
        time.sleep(5)

        wait.until(EC.presence_of_element_located((By.ID, "root")))
        print("CPLUS: 登錄後頁面加載完成", flush=True)
        return True
    except Exception as e:
        print(f"CPLUS 登錄失敗: {str(e)}", flush=True)
        driver.save_screenshot("cplus_login_error.png")
        save_page_source(driver, "cplus_login_error.html")
        print("CPLUS: 已保存登錄錯誤截圖和源碼: cplus_login_error.png, cplus_login_error.html", flush=True)
        return False

# 嘗試通過菜單導航
def navigate_via_menu(driver, wait, section):
    try:
        print(f"CPLUS: 嘗試通過菜單導航到 {section}...", flush=True)
        driver.get("https://cplus.hit.com.hk/frontpage/#/")
        time.sleep(2)
        menu_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@aria-label, 'menu') or contains(@class, 'MuiIconButton-root')]")))
        menu_button.click()
        print("CPLUS: 菜單按鈕點擊成功", flush=True)
        time.sleep(1)

        section_locators = {
            "ContainerMovementLog": ["//a[contains(text(), 'Container Movement Log')]", "//a[contains(text(), 'Enquiry')]"],
            "OnHandContainerList": ["//a[contains(text(), 'On Hand Container List')]", "//a[contains(text(), 'Enquiry')]"],
            "Housekeeping Reports": ["//a[contains(text(), 'Housekeeping Reports')]", "//a[contains(text(), 'Reports')]"]
        }

        for locator in section_locators[section]:
            try:
                section_link = wait.until(EC.element_to_be_clickable((By.XPATH, locator)))
                section_link.click()
                print(f"CPLUS: 通過菜單點擊 {section} 成功", flush=True)
                time.sleep(2)
                return True
            except TimeoutException:
                print(f"CPLUS: 菜單項 {locator} 未找到", flush=True)
        return False
    except Exception as e:
        print(f"CPLUS: 菜單導航失敗: {str(e)}", flush=True)
        return False

# Container Movement Log 下載
def process_container_movement_log(driver, wait, downloaded_files):
    max_retries = 3
    attempt = 0
    expected_files = 1
    new_files_count = 0

    while attempt < max_retries and new_files_count < expected_files:
        attempt += 1
        print(f"CPLUS Container Movement Log: 第 {attempt}/{max_retries} 次嘗試...", flush=True)
        try:
            print("CPLUS: 前往 Container Movement Log...", flush=True)
            driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
            time.sleep(2)
            if "404.html" in driver.current_url:
                print("CPLUS: 檢測到 404 錯誤，嘗試通過菜單導航...", flush=True)
                if not navigate_via_menu(driver, wait, "ContainerMovementLog"):
                    print("CPLUS: 菜單導航失敗，跳過此次嘗試", flush=True)
                    driver.save_screenshot(f"cplus_cml_404_attempt_{attempt}.png")
                    save_page_source(driver, f"cplus_cml_404_attempt_{attempt}.html")
                    continue

            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")))
            print("CPLUS: Container Movement Log 頁面加載完成", flush=True)

            print("CPLUS: 點擊 Search...", flush=True)
            initial_files = set(f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx')))
            search_button_locators = [
                "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button",
                "//button[contains(@class, 'MuiButtonBase-root') and .//span[text()='Search']]",
                "//button[.//span[text()='Search']]",
                "//button[contains(text(), 'Search')]"
            ]
            search_clicked = False
            for locator in search_button_locators:
                try:
                    search_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, locator)))
                    ActionChains(driver).move_to_element(search_button).click().perform()
                    print(f"CPLUS: Search 按鈕點擊成功 (使用定位: {locator})", flush=True)
                    search_clicked = True
                    break
                except TimeoutException:
                    print(f"CPLUS: Search 按鈕未找到 (定位: {locator})", flush=True)
            if not search_clicked:
                print("CPLUS: Container Movement Log Search 按鈕點擊失敗，當前 URL: ", driver.current_url, flush=True)
                driver.save_screenshot(f"cplus_cml_search_error_attempt_{attempt}.png")
                save_page_source(driver, f"cplus_cml_search_error_attempt_{attempt}.html")
                continue

            time.sleep(5)

            print("CPLUS: 點擊 Download...", flush=True)
            download_button_locators = [
                "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button",
                "//button[contains(text(), 'Download')]"
            ]
            download_clicked = False
            for locator in download_button_locators:
                try:
                    download_button = wait.until(EC.element_to_be_clickable((By.XPATH, locator)))
                    ActionChains(driver).move_to_element(download_button).click().perform()
                    print("CPLUS: Download 按鈕點擊成功", flush=True)
                    download_clicked = True
                    break
                except Exception as e:
                    print(f"CPLUS: Download 按鈕點擊失敗 (使用定位: {locator}): {str(e)}", flush=True)
            if not download_clicked:
                print("CPLUS: Container Movement Log Download 按鈕點擊失敗", flush=True)
                driver.save_screenshot(f"cplus_cml_download_error_attempt_{attempt}.png")
                save_page_source(driver, f"cplus_cml_download_error_attempt_{attempt}.html")
                continue

            new_files = wait_for_new_file(initial_files, timeout=15)
            if new_files:
                print(f"CPLUS: Container Movement Log 下載完成，檔案位於: {download_dir}", flush=True)
                for file in new_files:
                    print(f"CPLUS: 新下載檔案: {file}", flush=True)
                downloaded_files.update(new_files)
                new_files_count += len(new_files)
            else:
                print("CPLUS: Container Movement Log 未觸發新文件下載", flush=True)
                driver.save_screenshot(f"cplus_cml_no_download_attempt_{attempt}.png")
                save_page_source(driver, f"cplus_cml_no_download_attempt_{attempt}.html")
        except Exception as e:
            print(f"CPLUS Container Movement Log 錯誤: {str(e)}", flush=True)
            driver.save_screenshot(f"cplus_cml_error_attempt_{attempt}.png")
            save_page_source(driver, f"cplus_cml_error_attempt_{attempt}.html")
    return downloaded_files

# OnHandContainerList 下載
def process_onhand_container_list(driver, wait, downloaded_files):
    max_retries = 3
    attempt = 0
    expected_files = 1
    new_files_count = 0

    while attempt < max_retries and new_files_count < expected_files:
        attempt += 1
        print(f"CPLUS OnHandContainerList: 第 {attempt}/{max_retries} 次嘗試...", flush=True)
        try:
            print("CPLUS: 前往 OnHandContainerList 頁面...", flush=True)
            driver.get("https://cplus.hit.com.hk/app/#/enquiry/OnHandContainerList")
            time.sleep(2)
            if "404.html" in driver.current_url:
                print("CPLUS: 檢測到 404 錯誤，嘗試通過菜單導航...", flush=True)
                if not navigate_via_menu(driver, wait, "OnHandContainerList"):
                    print("CPLUS: 菜單導航失敗，跳過此次嘗試", flush=True)
                    driver.save_screenshot(f"cplus_ohcl_404_attempt_{attempt}.png")
                    save_page_source(driver, f"cplus_ohcl_404_attempt_{attempt}.html")
                    continue

            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
            print("CPLUS: OnHandContainerList 頁面加載完成", flush=True)

            print("CPLUS: 點擊 Search...", flush=True)
            initial_files = set(f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx')))
            search_button_onhand_locators = [
                "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]",
                "//button[contains(@class, 'MuiButtonBase-root') and .//span[text()='Search']]",
                "//button[.//span[text()='Search']]",
                "//button[contains(text(), 'Search')]"
            ]
            search_clicked = False
            for locator in search_button_onhand_locators:
                try:
                    search_button_onhand = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, locator)))
                    ActionChains(driver).move_to_element(search_button_onhand).click().perform()
                    print(f"CPLUS: Search 按鈕點擊成功 (使用定位: {locator})", flush=True)
                    search_clicked = True
                    break
                except TimeoutException:
                    print(f"CPLUS: Search 按鈕未找到 (定位: {locator})", flush=True)
            if not search_clicked:
                print("CPLUS: OnHandContainerList Search 按鈕點擊失敗，當前 URL: ", driver.current_url, flush=True)
                driver.save_screenshot(f"cplus_ohcl_search_error_attempt_{attempt}.png")
                save_page_source(driver, f"cplus_ohcl_search_error_attempt_{attempt}.html")
                continue

            time.sleep(5)

            print("CPLUS: 點擊 Export...", flush=True)
            export_button_locators = [
                "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button",
                "//button[contains(text(), 'Export')]"
            ]
            export_clicked = False
            for locator in export_button_locators:
                try:
                    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, locator)))
                    ActionChains(driver).move_to_element(export_button).click().perform()
                    print(f"CPLUS: Export 按鈕點擊成功 (使用定位: {locator})", flush=True)
                    export_clicked = True
                    break
                except Exception as e:
                    print(f"CPLUS: Export 按鈕點擊失敗 (使用定位: {locator}): {str(e)}", flush=True)
            if not export_clicked:
                print("CPLUS: OnHandContainerList Export 按鈕點擊失敗", flush=True)
                driver.save_screenshot(f"cplus_ohcl_export_error_attempt_{attempt}.png")
                save_page_source(driver, f"cplus_ohcl_export_error_attempt_{attempt}.html")
                continue

            print("CPLUS: 點擊 Export as CSV...", flush=True)
            csv_button_locators = [
                "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']",
                "//li[contains(text(), 'Export as CSV')]"
            ]
            csv_clicked = False
            for locator in csv_button_locators:
                try:
                    export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, locator)))
                    ActionChains(driver).move_to_element(export_csv_button).click().perform()
                    print(f"CPLUS: Export as CSV 按鈕點擊成功 (使用定位: {locator})", flush=True)
                    csv_clicked = True
                    break
                except Exception as e:
                    print(f"CPLUS: Export as CSV 按鈕點擊失敗 (使用定位: {locator}): {str(e)}", flush=True)
            if not csv_clicked:
                print("CPLUS: OnHandContainerList Export as CSV 按鈕點擊失敗", flush=True)
                driver.save_screenshot(f"cplus_ohcl_csv_error_attempt_{attempt}.png")
                save_page_source(driver, f"cplus_ohcl_csv_error_attempt_{attempt}.html")
                continue

            new_files = wait_for_new_file(initial_files, timeout=15)
            if new_files:
                print(f"CPLUS: OnHandContainerList 下載完成，檔案位於: {download_dir}", flush=True)
                for file in new_files:
                    print(f"CPLUS: 新下載檔案: {file}", flush=True)
                downloaded_files.update(new_files)
                new_files_count += len(new_files)
            else:
                print("CPLUS: OnHandContainerList 未觸發新文件下載", flush=True)
                driver.save_screenshot(f"cplus_ohcl_no_download_attempt_{attempt}.png")
                save_page_source(driver, f"cplus_ohcl_no_download_attempt_{attempt}.html")
        except Exception as e:
            print(f"CPLUS OnHandContainerList 錯誤: {str(e)}", flush=True)
            driver.save_screenshot(f"cplus_ohcl_error_attempt_{attempt}.png")
            save_page_source(driver, f"cplus_ohcl_error_attempt_{attempt}.html")
    return downloaded_files

# Housekeeping Reports 下載
def process_housekeeping_reports(driver, wait, downloaded_files):
    max_retries = 3
    attempt = 0
    expected_files = 6
    downloaded_reports = set()

    while attempt < max_retries and len(downloaded_files) < expected_files:
        attempt += 1
        print(f"CPLUS Housekeeping Reports: 第 {attempt}/{max_retries} 次嘗試...", flush=True)
        try:
            print("CPLUS: 前往 Housekeeping Reports 頁面...", flush=True)
            driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
            time.sleep(2)
            if "404.html" in driver.current_url:
                print("CPLUS: 檢測到 404 錯誤，嘗試通過菜單導航...", flush=True)
                if not navigate_via_menu(driver, wait, "Housekeeping Reports"):
                    print("CPLUS: 菜單導航失敗，跳過此次嘗試", flush=True)
                    driver.save_screenshot(f"cplus_hr_404_attempt_{attempt}.png")
                    save_page_source(driver, f"cplus_hr_404_attempt_{attempt}.html")
                    continue

            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
            print("CPLUS: Housekeeping Reports 頁面加載完成", flush=True)
            time.sleep(5)

            print("CPLUS: 檢查是否有日期選擇...", flush=True)
            try:
                date_fields = driver.find_elements(By.XPATH, "//input[contains(@id, 'date') or contains(@type, 'date')]")
                if date_fields:
                    for date_field in date_fields:
                        date_field.clear()
                        date_field.send_keys(datetime.now().strftime('%Y-%m-%d'))
                        print("CPLUS: 已填充日期字段", flush=True)
                        time.sleep(1)
            except Exception as e:
                print(f"CPLUS: 日期選擇失敗: {str(e)}", flush=True)

            print("CPLUS: 等待表格加載...", flush=True)
            try:
                wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr")))
                print("CPLUS: 表格加載完成", flush=True)
            except TimeoutException:
                print("CPLUS: 表格未加載，當前 URL: ", driver.current_url, ", 頁面源碼長度: ", len(driver.page_source), flush=True)
                driver.save_screenshot(f"cplus_hr_table_error_attempt_{attempt}.png")
                save_page_source(driver, f"cplus_hr_table_error_attempt_{attempt}.html")
                continue

            print("CPLUS: 定位所有 Excel 下載按鈕...", flush=True)
            initial_files = set(f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx')))
            locator_list = [
                "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]//svg[@viewBox='0 0 24 24']//path[@fill='#036e11']",
                "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)]",
                "//table//tbody//tr//td[4]//button[not(@disabled)]"
            ]
            excel_buttons = []
            selected_locator = None
            for locator in locator_list:
                excel_buttons = driver.find_elements(By.XPATH, locator)
                button_count = len(excel_buttons)
                print(f"CPLUS: 使用定位 {locator} 找到 {button_count} 個 Excel 下載按鈕", flush=True)
                if button_count > 0:
                    selected_locator = locator
                    break
            if button_count == 0:
                print("CPLUS: 未找到任何 Excel 下載按鈕，當前 URL: ", driver.current_url, ", 頁面源碼長度: ", len(driver.page_source), flush=True)
                driver.save_screenshot(f"cplus_hr_buttons_error_attempt_{attempt}.png")
                save_page_source(driver, f"cplus_hr_buttons_error_attempt_{attempt}.html")
                continue

            for idx in range(min(button_count, expected_files)):
                try:
                    button_locator = f"({selected_locator})[{idx+1}]"
                    button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, button_locator)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", button)
                    time.sleep(0.5)

                    try:
                        report_name = driver.find_element(By.XPATH, f"//table[contains(@class, 'MuiTable-root')]//tbody//tr[{idx+1}]//td[3]").text
                        print(f"CPLUS: 準備點擊第 {idx+1} 個 Excel 按鈕，報告名稱: {report_name}", flush=True)
                    except:
                        report_name = f"Unknown Report {idx+1}"
                        print(f"CPLUS: 無法獲取第 {idx+1} 個按鈕的報告名稱", flush=True)

                    if report_name in downloaded_reports:
                        print(f"CPLUS: 報告 {report_name} 已下載，跳過", flush=True)
                        continue

                    try:
                        is_disabled = driver.execute_script("return arguments[0].hasAttribute('disabled');", button)
                        button_class = driver.execute_script("return arguments[0].getAttribute('class');", button)
                        print(f"CPLUS: 第 {idx+1} 個按鈕狀態 - disabled: {is_disabled}, class: {button_class}", flush=True)
                    except:
                        print(f"CPLUS: 無法獲取第 {idx+1} 個按鈕狀態", flush=True)

                    clicked = False
                    for attempt_button in range(2):
                        try:
                            ActionChains(driver).move_to_element(button).pause(0.5).click().perform()
                            print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕點擊成功", flush=True)
                            clicked = True
                            break
                        except Exception as e:
                            print(f"CPLUS: 第 {idx+1} 個按鈕點擊嘗試 {attempt_button+1} 失敗: {str(e)}", flush=True)
                            time.sleep(0.5)
                    if not clicked:
                        print(f"CPLUS: 第 {idx+1} 個按鈕點擊失敗，重試 2 次後放棄", flush=True)
                        continue

                    try:
                        driver.execute_script("arguments[0].click();", button)
                        print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕 JavaScript 點擊成功", flush=True)
                    except Exception as js_e:
                        print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕 JavaScript 點擊失敗: {str(js_e)}", flush=True)

                    new_files = wait_for_new_file(initial_files, timeout=15)
                    if new_files:
                        print(f"CPLUS: 第 {idx+1} 個按鈕下載新文件: {', '.join(new_files)}", flush=True)
                        initial_files.update(new_files)
                        downloaded_files.update(new_files)
                        downloaded_reports.add(report_name)
                    else:
                        print(f"CPLUS: 第 {idx+1} 個按鈕未觸發新文件下載", flush=True)
                        driver.save_screenshot(f"cplus_hr_no_download_attempt_{attempt}_{idx+1}.png")
                        save_page_source(driver, f"cplus_hr_no_download_attempt_{attempt}_{idx+1}.html")

                    time.sleep(0.5)

                except ElementClickInterceptedException as e:
                    print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕點擊被攔截: {str(e)}", flush=True)
                    driver.save_screenshot(f"cplus_hr_button_intercepted_attempt_{attempt}_{idx+1}.png")
                    save_page_source(driver, f"cplus_hr_button_intercepted_attempt_{attempt}_{idx+1}.html")
                except TimeoutException as e:
                    print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕不可點擊: {str(e)}", flush=True)
                    driver.save_screenshot(f"cplus_hr_button_timeout_attempt_{attempt}_{idx+1}.png")
                    save_page_source(driver, f"cplus_hr_button_timeout_attempt_{attempt}_{idx+1}.html")
                except Exception as e:
                    print(f"CPLUS: 第 {idx+1} 個 Excel 下載按鈕處理失敗: {str(e)}", flush=True)
                    driver.save_screenshot(f"cplus_hr_button_error_attempt_{attempt}_{idx+1}.png")
                    save_page_source(driver, f"cplus_hr_button_error_attempt_{attempt}_{idx+1}.html")

        except Exception as e:
            print(f"CPLUS Housekeeping Reports 錯誤: {str(e)}", flush=True)
            driver.save_screenshot(f"cplus_hr_error_attempt_{attempt}.png")
            save_page_source(driver, f"cplus_hr_error_attempt_{attempt}.html")
    return downloaded_files

# CPLUS 主流程
def process_cplus():
    driver = None
    downloaded_files = set()
    try:
        driver = webdriver.Chrome(options=get_chrome_options())
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        wait = WebDriverWait(driver, 30)
        print("CPLUS WebDriver 初始化成功", flush=True)

        if not cplus_login(driver, wait):
            print("CPLUS: 登錄失敗，終止流程", flush=True)
            return downloaded_files

        downloaded_files = process_container_movement_log(driver, wait, downloaded_files)
        downloaded_files = process_onhand_container_list(driver, wait, downloaded_files)
        downloaded_files = process_housekeeping_reports(driver, wait, downloaded_files)

        return downloaded_files
    except Exception as e:
        print(f"CPLUS 總錯誤: {str(e)}", flush=True)
        driver.save_screenshot("cplus_total_error.png")
        save_page_source(driver, "cplus_total_error.html")
        return downloaded_files
    finally:
        if driver:
            try:
                print("CPLUS: 嘗試登出...", flush=True)
                logout_menu_locators = [
                    "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]",
                    "//button[contains(@aria-label, 'account')]"
                ]
                logout_option_locators = [
                    "//*[@id='menu-list-grow']/div[6]/li",
                    "//li[contains(text(), 'Logout')]"
                ]
                menu_clicked = False
                for menu_locator in logout_menu_locators:
                    try:
                        logout_menu_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, menu_locator)))
                        driver.execute_script("arguments[0].scrollIntoView(true);", logout_menu_button)
                        ActionChains(driver).move_to_element(logout_menu_button).click().perform()
                        print(f"CPLUS: 登出菜單按鈕點擊成功 (使用定位: {menu_locator})", flush=True)
                        menu_clicked = True
                        break
                    except TimeoutException:
                        print(f"CPLUS: 登出菜單按鈕未找到 (定位: {menu_locator})", flush=True)

                if menu_clicked:
                    option_clicked = False
                    for option_locator in logout_option_locators:
                        try:
                            logout_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, option_locator)))
                            driver.execute_script("arguments[0].scrollIntoView(true);", logout_option)
                            ActionChains(driver).move_to_element(logout_option).click().perform()
                            print(f"CPLUS: Logout 選項點擊成功 (使用定位: {option_locator})", flush=True)
                            option_clicked = True
                            break
                        except TimeoutException:
                            print(f"CPLUS: Logout 選項未找到 (定位: {option_locator})", flush=True)
                    if not option_clicked:
                        print("CPLUS: 無法點擊 Logout 選項，跳過登出", flush=True)
                else:
                    print("CPLUS: 無法打開登出菜單，跳過登出", flush=True)
                time.sleep(2)
            except Exception as logout_error:
                print(f"CPLUS: 登出失敗: {str(logout_error)}", flush=True)

            driver.quit()
            print("CPLUS WebDriver 關閉", flush=True)

# Barge 操作
def process_barge():
    driver = None
    downloaded_files = set()
    max_retries = 3
    attempt = 0
    expected_files = 1

    while attempt < max_retries and len(downloaded_files) < expected_files:
        attempt += 1
        print(f"Barge: 第 {attempt}/{max_retries} 次嘗試...", flush=True)
        try:
            if not driver:
                driver = webdriver.Chrome(options=get_chrome_options())
                driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
                print("Barge WebDriver 初始化成功", flush=True)
                wait = WebDriverWait(driver, 30)

                print("Barge: 嘗試打開網站 https://barge.oneport.com/downloadReport...", flush=True)
                driver.get("https://barge.oneport.com/downloadReport")
                print(f"Barge: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)
                time.sleep(3)

                if "login" in driver.current_url:
                    print("Barge: 檢測到登錄頁面，開始輸入登錄信息...", flush=True)
                    print("Barge: 輸入 COMPANY ID...", flush=True)
                    company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='mat-input-0']")))
                    company_id_field.clear()
                    company_id_field.send_keys("CKL")
                    print("Barge: COMPANY ID 輸入完成", flush=True)
                    time.sleep(1)

                    print("Barge: 輸入 USER ID...", flush=True)
                    user_id_field_barge = driver.find_element(By.XPATH, "//*[@id='mat-input-1']")
                    user_id_field_barge.clear()
                    user_id_field_barge.send_keys("barge")
                    print("Barge: USER ID 輸入完成", flush=True)
                    time.sleep(1)

                    print("Barge: 輸入 PW...", flush=True)
                    password_field_barge = driver.find_element(By.XPATH, "//*[@id='mat-input-2']")
                    password_field_barge.clear()
                    password_field_barge.send_keys(os.environ.get('BARGE_PASSWORD', '123456'))
                    print("Barge: PW 輸入完成", flush=True)
                    time.sleep(1)

                    print("Barge: 點擊 LOGIN 按鈕...", flush=True)
                    login_button_barge = driver.find_element(By.XPATH, "//*[@id='login-form-container']/app-login-form/form/div/button")
                    ActionChains(driver).move_to_element(login_button_barge).click().perform()
                    print("Barge: LOGIN 按鈕點擊成功", flush=True)
                    time.sleep(5)

                    wait.until(EC.presence_of_element_located((By.ID, "content-mount")))
                    print("Barge: 登錄後頁面加載完成", flush=True)

                print("Barge: 檢查是否有日期選擇...", flush=True)
                try:
                    date_fields = driver.find_elements(By.XPATH, "//input[contains(@id, 'date') or contains(@type, 'date')]")
                    if date_fields:
                        for date_field in date_fields:
                            date_field.clear()
                            date_field.send_keys(datetime.now().strftime('%Y-%m-%d'))
                            print("Barge: 已填充日期字段", flush=True)
                            time.sleep(1)
                except Exception as e:
                    print(f"Barge: 日期選擇失敗: {str(e)}", flush=True)

                print("Barge: 選擇 Report Type...", flush=True)
                report_type_select_locators = [
                    "//mat-select[contains(@aria-label, 'Report Type') or contains(@id, 'mat-select')]",
                    "//select[contains(@id, 'mat-select') or contains(@class, 'mat-select')]",
                    "//*[contains(text(), 'Report Type')]/following-sibling::mat-select"
                ]
                type_selected = False
                for locator in report_type_select_locators:
                    try:
                        report_type_select = wait.until(EC.element_to_be_clickable((By.XPATH, locator)))
                        ActionChains(driver).move_to_element(report_type_select).click().perform()
                        print(f"Barge: Report Type 選擇開始 (使用定位: {locator})", flush=True)
                        type_selected = True
                        break
                    except TimeoutException:
                        print(f"Barge: Report Type 未找到 (定位: {locator})", flush=True)
                if not type_selected:
                    print("Barge: Report Type 選擇失敗，嘗試直接選擇 Container Detail...", flush=True)

                time.sleep(2)

                print("Barge: 點擊 Container Detail...", flush=True)
                container_detail_locators = [
                    "//mat-option[contains(text(), 'Container Detail')]",
                    "//*[contains(text(), 'Container Detail') and @role='option']",
                    "//*[contains(text(), 'Container Detail')]"
                ]
                detail_clicked = False
                for locator in container_detail_locators:
                    try:
                        container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, locator)))
                        ActionChains(driver).move_to_element(container_detail_option).click().perform()
                        print(f"Barge: Container Detail 點擊成功 (使用定位: {locator})", flush=True)
                        detail_clicked = True
                        break
                    except TimeoutException:
                        print(f"Barge: Container Detail 選項未找到 (定位: {locator})", flush=True)
                if not detail_clicked:
                    print("Barge: Container Detail 點擊失敗", flush=True)
                    driver.save_screenshot(f"barge_container_detail_error_attempt_{attempt}.png")
                    save_page_source(driver, f"barge_container_detail_error_attempt_{attempt}.html")
                    continue

                time.sleep(2)

                print("Barge: 點擊 Download...", flush=True)
                initial_files = set(f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx')))
                download_button_barge_locators = [
                    "//*[@id='content-mount']/app-download-report/div[2]/div/form/div[2]/button",
                    "//button[contains(text(), 'Download')]"
                ]
                download_clicked = False
                for locator in download_button_barge_locators:
                    try:
                        download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, locator)))
                        ActionChains(driver).move_to_element(download_button_barge).click().perform()
                        print("Barge: Download 按鈕點擊成功", flush=True)
                        download_clicked = True
                        break
                    except Exception as e:
                        print(f"Barge: Download 按鈕點擊失敗 (使用定位: {locator}): {str(e)}", flush=True)
                if not download_clicked:
                    print("Barge: Container Detail Download 按鈕點擊失敗", flush=True)
                    driver.save_screenshot(f"barge_download_error_attempt_{attempt}.png")
                    save_page_source(driver, f"barge_download_error_attempt_{attempt}.html")
                    continue

                new_files = wait_for_new_file(initial_files, timeout=15)
                if new_files:
                    print(f"Barge: Container Detail 下載完成，檔案位於: {download_dir}", flush=True)
                    for file in new_files:
                        print(f"Barge: 新下載檔案: {file}", flush=True)
                    downloaded_files.update(new_files)
                    break
                else:
                    print("Barge: Container Detail 未觸發新文件下載", flush=True)
                    driver.save_screenshot(f"barge_no_download_attempt_{attempt}.png")
                    save_page_source(driver, f"barge_no_download_attempt_{attempt}.html")
                    continue

        except Exception as e:
            print(f"Barge 錯誤: {str(e)}", flush=True)
            if driver:
                driver.quit()
                driver = None
            time.sleep(10)
            continue

    if driver:
        try:
            print("Barge: 嘗試登出...", flush=True)
            logout_toolbar_locators = [
                "//*[@id='main-toolbar']/button[4]/span[1]",
                "//button[contains(@aria-label, 'account')]"
            ]
            logout_option_locators = [
                "//*[@id='mat-menu-panel-11']/div/button/span",
                "//button[contains(text(), 'Logout')]"
            ]
            toolbar_clicked = False
            for toolbar_locator in logout_toolbar_locators:
                try:
                    logout_toolbar_barge = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, toolbar_locator)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar_barge)
                    ActionChains(driver).move_to_element(logout_toolbar_barge).click().perform()
                    print(f"Barge: 工具欄點擊成功 (使用定位: {toolbar_locator})", flush=True)
                    toolbar_clicked = True
                    break
                except TimeoutException:
                    print(f"Barge: 工具欄按鈕未找到 (定位: {toolbar_locator})", flush=True)

            if toolbar_clicked:
                option_clicked = False
                for option_locator in logout_option_locators:
                    try:
                        logout_button_barge = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, option_locator)))
                        driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
                        ActionChains(driver).move_to_element(logout_button_barge).click().perform()
                        print(f"Barge: Logout 選項點擊成功 (使用定位: {option_locator})", flush=True)
                        option_clicked = True
                        break
                    except TimeoutException:
                        print(f"Barge: Logout 選項未找到 (定位: {option_locator})", flush=True)
                if not option_clicked:
                    print("Barge: 無法點擊 Logout 選項，跳過登出", flush=True)
            else:
                print("Barge: 無法打開工具欄，跳過登出", flush=True)
            time.sleep(2)
        except Exception as e:
            print(f"Barge: 登出失敗: {str(e)}", flush=True)

        driver.quit()
        print("Barge WebDriver 關閉", flush=True)

    return downloaded_files

# 主函數
def main():
    setup_environment()
    clear_download_dir()

    cplus_files = [set()]
    barge_files = [set()]

    def run_cplus():
        cplus_files[0] = process_cplus()

    def run_barge():
        barge_files[0] = process_barge()

    t1 = threading.Thread(target=run_cplus)
    t2 = threading.Thread(target=run_barge)

    t1.start()
    t2.start()

    t1.join()
    t2.join()

    all_downloaded_files = cplus_files[0].union(barge_files[0]) if cplus_files[0] and barge_files[0] else set()
    downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
    
    # 過濾重複文件（保留最新版本）
    unique_files = {}
    for file in downloaded_files:
        base_name = file.split(' (')[0]
        if base_name not in unique_files or ' (' not in file:
            unique_files[base_name] = file
        elif ' (' in file:
            current_num = int(file.split(' (')[1].split(')')[0]) if ' (' in file else 0
            existing_num = int(unique_files[base_name].split(' (')[1].split(')')[0]) if ' (' in unique_files[base_name] else 0
            if current_num > existing_num:
                unique_files[base_name] = file

    final_files = list(unique_files.values())
    print(f"\n檢查所有下載文件: {len(final_files)}/{EXPECTED_TOTAL_FILES}", flush=True)

    if len(final_files) >= EXPECTED_TOTAL_FILES:
        print(f"所有下載完成，檔案位於: {download_dir}", flush=True)
        for file in final_files:
            print(f"找到檔案: {file}", flush=True)

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

            for file in final_files:
                file_path = os.path.join(download_dir, file)
                if os.path.getsize(file_path) > 0:
                    attachment = MIMEBase('application', 'octet-stream')
                    attachment.set_payload(open(file_path, 'rb').read())
                    encoders.encode_base64(attachment)
                    attachment.add_header('Content-Disposition', f'attachment; filename={file}')
                    msg.attach(attachment)
                else:
                    print(f"警告: 檔案 {file} 為空，跳過附件", flush=True)

            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
            server.quit()
            print("郵件發送成功!", flush=True)
        except Exception as e:
            print(f"郵件發送失敗: {str(e)}", flush=True)
    else:
        print(f"下載文件數量不足（{len(final_files)}/{EXPECTED_TOTAL_FILES}），無法發送郵件", flush=True)

    print("腳本完成", flush=True)

if __name__ == "__main__":
    main()
