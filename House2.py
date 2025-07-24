```python
import os
import time
import subprocess
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

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

def check_chromium_compatibility():
    try:
        result = subprocess.run(['chromium-browser', '--version'], capture_output=True, text=True, check=True)
        chromium_version = result.stdout.strip().split()[1].split('.')[0]
        print(f"Chromium 版本: {chromium_version}", flush=True)

        chromedriver_path = ChromeDriverManager().install()
        result = subprocess.run([chromedriver_path, '--version'], capture_output=True, text=True, check=True)
        chromedriver_version = result.stdout.strip().split()[1].split('.')[0]
        print(f"Chromedriver 版本: {chromedriver_version}", flush=True)

        if chromium_version != chromedriver_version:
            print(f"警告: Chromium ({chromium_version}) 與 Chromedriver ({chromedriver_version}) 版本不完全匹配，可能導致兼容性問題", flush=True)
        else:
            print("Chromium 和 Chromedriver 版本兼容", flush=True)
        return chromedriver_path
    except subprocess.CalledProcessError as e:
        print(f"檢查版本失敗: {e}", flush=True)
        raise

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
    chrome_options.binary_location = '/usr/bin/chromium-browser'
    return chrome_options

def login_cplus(driver, company_code, user_id, password):
    wait = WebDriverWait(driver, 30)
    try:
        driver.get("https://cplus.hit.com.hk/frontpage/#/")
        print(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)

        login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
        login_button_pre.click()
        print("CPLUS: 登錄前按鈕點擊成功", flush=True)

        company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
        company_code_field.send_keys(company_code)
        print("CPLUS: COMPANY CODE 輸入完成", flush=True)

        user_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='userId']")))
        user_id_field.send_keys(user_id)
        print("CPLUS: USER ID 輸入完成", flush=True)

        password_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='passwd']")))
        password_field.send_keys(password)
        print("CPLUS: PASSWORD 輸入完成", flush=True)

        login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")))
        login_button.click()
        print("CPLUS: LOGIN 按鈕點擊成功", flush=True)

        # 檢查登錄是否成功
        wait.until(EC.url_changes("https://cplus.hit.com.hk/frontpage/#/"))
        if "login" in driver.current_url.lower() or "404" in driver.current_url:
            raise Exception(f"登錄失敗，當前 URL: {driver.current_url}")
    except TimeoutException as e:
        print(f"CPLUS: 登錄失敗: {str(e)}", flush=True)
        print(f"當前 URL: {driver.current_url}", flush=True)
        print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
        raise

def navigate_to_housekeep_report(driver):
    wait = WebDriverWait(driver, 30)
    try:
        print("Download Housekeep: 嘗試導航到 Housekeep Report 頁面...", flush=True)
        # 嘗試直接訪問 URL
        driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
        
        # 檢查是否跳轉到 404
        if "404.html" in driver.current_url:
            print(f"Download Housekeep: 訪問失敗，跳轉到 404，當前 URL: {driver.current_url}", flush=True)
            print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
            # 嘗試通過導航點擊進入
            print("Download Housekeep: 嘗試通過導航點擊進入報告頁面...", flush=True)
            report_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Report')] | //li[contains(text(), 'Report')]")))
            driver.execute_script("arguments[0].click();", report_menu)
            housekeep_report = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Housekeep Report')] | //li[contains(text(), 'Housekeep Report')]")))
            driver.execute_script("arguments[0].click();", housekeep_report)
            wait.until(EC.url_contains("housekeepReport"))
        
        print("Download Housekeep: Housekeep Report 頁面加載完成", flush=True)
        print(f"當前 URL: {driver.current_url}", flush=True)
    except TimeoutException as e:
        print(f"Download Housekeep: 導航到報告頁面失敗: {str(e)}", flush=True)
        print(f"當前 URL: {driver.current_url}", flush=True)
        print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
        with open("page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise

def download_housekeep_report(driver):
    wait = WebDriverWait(driver, 30)
    try:
        navigate_to_housekeep_report(driver)

        # 等待表單容器加載
        wait.until(EC.presence_of_element_located((By.XPATH, "//form")))
        print("Download Housekeep: 表單容器加載完成", flush=True)

        today = datetime.now().strftime("%d/%m/%Y")  # 格式：24/07/2025
        print(f"Download Housekeep: 檢查日期，今日為 {today}", flush=True)

        # 檢查 From 輸入框
        try:
            from_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='from'] | //input[@id='from']")))
            from_value = from_field.get_attribute("value")
            if from_value != today:
                print(f"Download Housekeep: From 日期 ({from_value}) 不為今日，設置為 {today}", flush=True)
                from_field.clear()
                from_field.send_keys(today)
            else:
                print("Download Housekeep: From 日期已正確", flush=True)
        except TimeoutException:
            print("Download Housekeep: 未找到 From 輸入框", flush=True)
            print(f"當前 URL: {driver.current_url}", flush=True)
            print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
            with open("page_source.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)

        # 檢查 To 輸入框
        try:
            to_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='to'] | //input[@id='to']")))
            to_value = to_field.get_attribute("value")
            if to_value != today:
                print(f"Download Housekeep: To 日期 ({to_value}) 不為今日，設置為 {today}", flush=True)
                to_field.clear()
                to_field.send_keys(today)
            else:
                print("Download Housekeep: To 日期已正確", flush=True)
        except TimeoutException:
            print("Download Housekeep: 未找到 To 輸入框", flush=True)
            print(f"當前 URL: {driver.current_url}", flush=True)
            print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
            with open("page_source.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)

        print("Download Housekeep: 查找並點擊所有 Email Excel checkbox...", flush=True)
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//tbody/tr")))
            checkboxes = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody//input[@type='checkbox']")))
            print(f"Download Housekeep: 找到 {len(checkboxes)} 個 Email Excel checkbox", flush=True)

            if len(checkboxes) < 3 or len(checkboxes) > 6:
                print(f"Download Housekeep: Checkbox 數量 {len(checkboxes)} 不在預期範圍 (3-6)，請檢查", flush=True)

            for index, checkbox in enumerate(checkboxes, 1):
                if checkbox.is_enabled() and checkbox.is_displayed() and not checkbox.is_selected():
                    driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                    driver.execute_script("arguments[0].click();", checkbox)
                    print(f"Download Housekeep: Checkbox {index} 點擊成功", flush=True)
        except TimeoutException:
            print("Download Housekeep: 未找到 Email Excel checkbox，嘗試備用定位...", flush=True)
            checkboxes = driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
            if checkboxes:
                for index, checkbox in enumerate(checkboxes, 1):
                    if checkbox.is_enabled() and checkbox.is_displayed() and not checkbox.is_selected():
                        driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                        driver.execute_script("arguments[0].click();", checkbox)
                        print(f"Download Housekeep: Checkbox {index} 點擊成功 (備用定位)", flush=True)
            else:
                print("Download Housekeep: 備用定位也未找到 checkbox，任務中止", flush=True)
                print(f"當前 URL: {driver.current_url}", flush=True)
                print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
                with open("page_source.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                return

        print("Download Housekeep: 點擊 Email 按鈕...", flush=True)
        try:
            email_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/form/div[1]/div[8]/div/div/div[1]/div[1]/div[1]/button")))
            driver.execute_script("arguments[0].scrollIntoView(true);", email_button)
            driver.execute_script("arguments[0].click();", email_button)
            print("Download Housekeep: Email 按鈕點擊成功", flush=True)

            target_email = os.environ.get('TARGET_EMAIL', 'paklun@ckline.com.hk')
            print("Download Housekeep: 輸入目標 Email 地址...", flush=True)
            email_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='to']")))
            email_field.clear()
            email_field.send_keys(target_email)
            print("Download Housekeep: Email 地址輸入完成", flush=True)

            current_time = datetime.now().strftime("%m:%d %H:%M")
            print(f"Download Housekeep: 輸入內文，格式為 {current_time}", flush=True)
            body_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='body']")))
            body_field.clear()
            body_field.send_keys(current_time)
            print("Download Housekeep: 內文輸入完成", flush=True)

            print("Download Housekeep: 點擊 Confirm 按鈕...", flush=True)
            confirm_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='EmailDialog']/div[3]/div/form/div[3]/button[1]")))
            driver.execute_script("arguments[0].click();", confirm_button)
            print("Download Housekeep: Confirm 按鈕點擊成功", flush=True)
        except TimeoutException as e:
            print(f"Download Housekeep: Email 處理失敗: {str(e)}", flush=True)
            print(f"當前 URL: {driver.current_url}", flush=True)
            print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
            with open("page_source.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            return
    except Exception as e:
        print(f"Download Housekeep: 錯誤: {str(e)}", flush=True)
        print(f"當前 URL: {driver.current_url}", flush=True)
        print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
        with open("page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        return

def main():
    setup_environment()
    chromedriver_path = check_chromium_compatibility()
    driver = None
    try:
        driver = webdriver.Chrome(service=Service(chromedriver_path), options=get_chrome_options())
        driver.set_page_load_timeout(30)
        print("CPLUS WebDriver 初始化成功", flush=True)

        company_code = os.environ.get('COMPANY_CODE', 'CKL')
        user_id = os.environ.get('USER_ID', 'KEN')
        password = os.environ.get('SITE_PASSWORD')
        if not password:
            raise ValueError("環境變量 SITE_PASSWORD 未設置")

        login_cplus(driver, company_code, user_id, password)
        download_housekeep_report(driver)

    except Exception as e:
        print(f"Download Housekeep 錯誤: {str(e)}", flush=True)
    finally:
        if driver:
            try:
                print("Download Housekeep: 嘗試登出...", flush=True)
                logout_menu_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
                driver.execute_script("arguments[0].click();", logout_menu_button)
                print("Download Housekeep: 登錄按鈕點擊成功", flush=True)

                try:
                    logout_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Logout')] | //a[contains(text(), 'Logout')] | //li[contains(text(), 'Logout')]")))
                    driver.execute_script("arguments[0].click();", logout_option)
                    print("Download Housekeep: Logout 選項點擊成功", flush=True)
                except TimeoutException:
                    print("Download Housekeep: 登出選項未找到，跳過登出", flush=True)
            except Exception as logout_error:
                print(f"Download Housekeep: 登出失敗: {str(logout_error)}", flush=True)
            driver.quit()
            print("Download Housekeep WebDriver 關閉", flush=True)

if __name__ == "__main__":
    main()
    print("Download Housekeep 腳本完成", flush=True)
