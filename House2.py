import os
import time
import subprocess
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
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
    wait = WebDriverWait(driver, 60)
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

        # 等待頁面加載完成
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
            time.sleep(3)  # 額外等待確保頁面穩定
        except TimeoutException:
            print("CPLUS: 頁面加載超時，嘗試刷新頁面...", flush=True)
            driver.refresh()
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
            time.sleep(3)

        final_url = driver.current_url
        print(f"CPLUS: 登錄後 URL: {final_url}", flush=True)
        
        # 檢查是否有錯誤提示
        try:
            error_message = driver.find_element(By.XPATH, "//*[contains(text(), 'error') or contains(text(), '失敗') or contains(text(), 'failed') or contains(text(), 'invalid')]")
            print(f"CPLUS: 檢測到錯誤提示: {error_message.text}", flush=True)
            with open("page_source.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            raise Exception(f"登錄失敗，錯誤提示: {error_message.text}")
        except:
            print("CPLUS: 未檢測到錯誤提示", flush=True)

        # 檢查是否需要兩步驗證
        try:
            two_factor_input = driver.find_element(By.XPATH, "//input[@type='text' or @type='password' or contains(@id, 'otp') or contains(@id, '2fa')]")
            print("CPLUS: 檢測到可能的兩步驗證輸入框，需手動處理", flush=True)
            with open("page_source.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            raise Exception("檢測到兩步驗證，請手動輸入驗證碼")
        except:
            print("CPLUS: 未檢測到兩步驗證輸入框", flush=True)

        if "app" not in final_url.lower():
            raise Exception(f"登錄失敗，當前 URL: {final_url}")
    except TimeoutException as e:
        print(f"CPLUS: 登錄失敗: {str(e)}", flush=True)
        print(f"當前 URL: {driver.current_url}", flush=True)
        print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
        with open("page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise
    except Exception as e:
        print(f"CPLUS: 登錄失敗: {str(e)}", flush=True)
        print(f"當前 URL: {driver.current_url}", flush=True)
        print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
        with open("page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise

def navigate_to_housekeep_report(driver):
    wait = WebDriverWait(driver, 60)
    try:
        print("Download Housekeep: 嘗試導航到 Housekeep Report 頁面...", flush=True)
        driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
        
        if "404.html" in driver.current_url:
            print(f"Download Housekeep: 訪問失敗，跳轉到 404，當前 URL: {driver.current_url}", flush=True)
            print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
            with open("page_source.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            raise Exception("訪問 Housekeep Report 頁面失敗，跳轉到 404")
        
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
    wait = WebDriverWait(driver, 60)
    try:
        navigate_to_housekeep_report(driver)

        wait.until(EC.presence_of_element_located((By.XPATH, "//form")))
        print("Download Housekeep: 表單容器加載完成", flush=True)

        today = datetime.now().strftime("%d/%m/%Y")
        print(f"Download Housekeep: 檢查日期，今日為 {today}", flush=True)

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
            wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr")))
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            any_checked = False
            for index in range(1, 7):  # 假設最多 6 個 checkbox
                try:
                    checkboxes = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr/td[6]//input[@type='checkbox']")))
                    if index > len(checkboxes):
                        break
                    checkbox = checkboxes[index - 1]
                    is_enabled = checkbox.is_enabled()
                    is_selected = driver.execute_script("return arguments[0].checked;", checkbox)
                    print(f"Download Housekeep: Checkbox {index} 狀態 - 啟用: {is_enabled}, 已選中: {is_selected}", flush=True)
                    
                    if is_enabled and not is_selected:
                        driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                        driver.execute_script("arguments[0].click();", checkbox)
                        time.sleep(1)
                        is_selected_after = driver.execute_script("return arguments[0].checked;", checkbox)
                        if is_selected_after:
                            print(f"Download Housekeep: Checkbox {index} 點擊成功", flush=True)
                            any_checked = True
                        else:
                            print(f"Download Housekeep: Checkbox {index} 點擊失敗，未選中，嘗試設置 checked 屬性", flush=True)
                            driver.execute_script("arguments[0].checked = true;", checkbox)
                            is_selected_after = driver.execute_script("return arguments[0].checked;", checkbox)
                            if is_selected_after:
                                print(f"Download Housekeep: Checkbox {index} 設置成功", flush=True)
                                any_checked = True
                            else:
                                print(f"Download Housekeep: Checkbox {index} 設置失敗，未選中", flush=True)
                    else:
                        print(f"Download Housekeep: Checkbox {index} 已選中或不可點擊，跳過", flush=True)
                        if is_selected:
                            any_checked = True
                except StaleElementReferenceException:
                    print(f"Download Housekeep: Checkbox {index} 遇到 StaleElementReferenceException，重新查找...", flush=True)
                    checkboxes = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr/td[6]//input[@type='checkbox']")))
                    if index > len(checkboxes):
                        break
                    checkbox = checkboxes[index - 1]
                    is_enabled = checkbox.is_enabled()
                    is_selected = driver.execute_script("return arguments[0].checked;", checkbox)
                    print(f"Download Housekeep: Checkbox {index} 狀態 (重新查找) - 啟用: {is_enabled}, 已選中: {is_selected}", flush=True)
                    
                    if is_enabled and not is_selected:
                        driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                        driver.execute_script("arguments[0].click();", checkbox)
                        time.sleep(1)
                        is_selected_after = driver.execute_script("return arguments[0].checked;", checkbox)
                        if is_selected_after:
                            print(f"Download Housekeep: Checkbox {index} 點擊成功 (重新查找)", flush=True)
                            any_checked = True
                        else:
                            print(f"Download Housekeep: Checkbox {index} 點擊失敗 (重新查找)，嘗試設置 checked 屬性", flush=True)
                            driver.execute_script("arguments[0].checked = true;", checkbox)
                            is_selected_after = driver.execute_script("return arguments[0].checked;", checkbox)
                            if is_selected_after:
                                print(f"Download Housekeep: Checkbox {index} 設置成功 (重新查找)", flush=True)
                                any_checked = True
                            else:
                                print(f"Download Housekeep: Checkbox {index} 設置失敗 (重新查找)，未選中", flush=True)
                    else:
                        print(f"Download Housekeep: Checkbox {index} 已選中或不可點擊 (重新查找)，跳過", flush=True)
                        if is_selected:
                            any_checked = True

            if not any_checked:
                print("Download Housekeep: 無任何 Checkbox 被選中，可能影響 Email 按鈕", flush=True)
                with open("page_source.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                
        except TimeoutException:
            print("Download Housekeep: 未找到 Email Excel checkbox，嘗試備用定位...", flush=True)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            checkboxes = driver.find_elements(By.CSS_SELECTOR, "td:nth-child(6) input[type='checkbox']")
            if checkboxes:
                any_checked = False
                for index, checkbox in enumerate(checkboxes, 1):
                    try:
                        is_enabled = checkbox.is_enabled()
                        is_selected = driver.execute_script("return arguments[0].checked;", checkbox)
                        print(f"Download Housekeep: Checkbox {index} 狀態 (備用定位) - 啟用: {is_enabled}, 已選中: {is_selected}", flush=True)
                        
                        if is_enabled and not is_selected:
                            driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                            driver.execute_script("arguments[0].click();", checkbox)
                            time.sleep(1)
                            is_selected_after = driver.execute_script("return arguments[0].checked;", checkbox)
                            if is_selected_after:
                                print(f"Download Housekeep: Checkbox {index} 點擊成功 (備用定位)", flush=True)
                                any_checked = True
                            else:
                                print(f"Download Housekeep: Checkbox {index} 點擊失敗 (備用定位)，嘗試設置 checked 屬性", flush=True)
                                driver.execute_script("arguments[0].checked = true;", checkbox)
                                is_selected_after = driver.execute_script("return arguments[0].checked;", checkbox)
                                if is_selected_after:
                                    print(f"Download Housekeep: Checkbox {index} 設置成功 (備用定位)", flush=True)
                                    any_checked = True
                                else:
                                    print(f"Download Housekeep: Checkbox {index} 設置失敗 (備用定位)，未選中", flush=True)
                        else:
                            print(f"Download Housekeep: Checkbox {index} 已選中或不可點擊 (備用定位)，跳過", flush=True)
                            if is_selected:
                                any_checked = True
