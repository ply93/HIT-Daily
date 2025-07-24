import os
import time
import subprocess
import shutil
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

def save_screenshot(driver, screenshot_dir):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    screenshot_path = os.path.join(screenshot_dir, f"screenshot_{timestamp}.png")
    driver.save_screenshot(screenshot_path)
    print(f"已保存截圖: {screenshot_path}", flush=True)
    return screenshot_path

def login_cplus(driver, company_code, user_id, password, screenshot_dir):
    wait = WebDriverWait(driver, 20)
    try:
        print("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/", flush=True)
        driver.get("https://cplus.hit.com.hk/frontpage/#/")
        print(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)
        time.sleep(2)

        print("CPLUS: 點擊登錄前按鈕...", flush=True)
        login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
        login_button_pre.click()
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
        login_button.click()
        print("CPLUS: LOGIN 按鈕點擊成功", flush=True)
        time.sleep(5)

        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
            print("CPLUS: 檢測到 root 元素，頁面加載完成", flush=True)
        except TimeoutException:
            print("CPLUS: 頁面加載超時，嘗試刷新頁面...", flush=True)
            driver.refresh()
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
            time.sleep(5)

        final_url = driver.current_url
        print(f"CPLUS: 登錄後 URL: {final_url}", flush=True)
        
        try:
            error_message = driver.find_element(By.XPATH, "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'error') or contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'failed') or contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'invalid')]")
            print(f"CPLUS: 檢測到錯誤提示: {error_message.text}", flush=True)
            with open("page_source.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            save_screenshot(driver, screenshot_dir)
            raise Exception(f"登錄失敗，錯誤提示: {error_message.text}")
        except:
            print("CPLUS: 未檢測到錯誤提示", flush=True)

        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
            print("CPLUS: 檢測到登出按鈕，假設登錄成功", flush=True)
        except TimeoutException:
            print("CPLUS: 未檢測到登出按鈕，登錄可能失敗", flush=True)
            with open("page_source.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            save_screenshot(driver, screenshot_dir)
            raise Exception(f"登錄失敗，當前 URL: {final_url}")

    except Exception as e:
        print(f"CPLUS: 登錄失敗: {str(e)}", flush=True)
        print(f"當前 URL: {driver.current_url}", flush=True)
        print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
        with open("page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        save_screenshot(driver, screenshot_dir)
        raise

def navigate_to_housekeep_report(driver, screenshot_dir):
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
            save_screenshot(driver, screenshot_dir)
            raise Exception("訪問 Housekeep Report 頁面失敗，跳轉到 404")
        
        print("Download Housekeep: Housekeep Report 頁面加載完成", flush=True)
        print(f"當前 URL: {driver.current_url}", flush=True)
    except TimeoutException as e:
        print(f"Download Housekeep: 導航到報告頁面失敗: {str(e)}", flush=True)
        print(f"當前 URL: {driver.current_url}", flush=True)
        print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
        with open("page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        save_screenshot(driver, screenshot_dir)
        raise

def download_housekeep_report(driver, screenshot_dir):
    wait = WebDriverWait(driver, 90)
    try:
        navigate_to_housekeep_report(driver, screenshot_dir)

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
            save_screenshot(driver, screenshot_dir)
            raise

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
            save_screenshot(driver, screenshot_dir)
            raise

        print("Download Housekeep: 查找並點擊所有 Email Excel checkbox...", flush=True)
        try:
            wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr")))
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(30)  # 等待表格穩定
            any_checked = False
            for index in range(1, 7):
                try:
                    checkboxes = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr/td[6]//input[@type='checkbox']")))
                    if index > len(checkboxes):
                        break
                    checkbox = checkboxes[index - 1]
                    is_enabled = checkbox.is_enabled()
                    is_selected = driver.execute_script("return arguments[0].checked;", checkbox)
                    print(f"Download Housekeep: Checkbox {index} 狀態 - 啟用: {is_enabled}, 已選中: {is_selected}", flush=True)
                    
                    if is_enabled:
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
                        print(f"Download Housekeep: Checkbox {index} 不可點擊，跳過", flush=True)
                except StaleElementReferenceException:
                    print(f"Download Housekeep: Checkbox {index} 遇到 StaleElementReferenceException，重新查找...", flush=True)
                    checkboxes = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr/td[6]//input[@type='checkbox']")))
                    if index > len(checkboxes):
                        break
                    checkbox = checkboxes[index - 1]
                    is_enabled = checkbox.is_enabled()
                    is_selected = driver.execute_script("return arguments[0].checked;", checkbox)
                    print(f"Download Housekeep: Checkbox {index} 狀態 (重新查找) - 啟用: {is_enabled}, 已選中: {is_selected}", flush=True)
                    
                    if is_enabled:
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
                        print(f"Download Housekeep: Checkbox {index} 不可點擊 (重新查找)，跳過", flush=True)

            if not any_checked:
                print("Download Housekeep: 無任何 Checkbox 被選中，可能影響 Email 按鈕", flush=True)
                with open("page_source.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                save_screenshot(driver, screenshot_dir)
                return
                
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
                        
                        if is_enabled:
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
                            print(f"Download Housekeep: Checkbox {index} 不可點擊 (備用定位)，跳過", flush=True)
                    except StaleElementReferenceException:
                        print(f"Download Housekeep: Checkbox {index} 遇到 StaleElementReferenceException (備用定位)，任務中止", flush=True)
                        with open("page_source.html", "w", encoding="utf-8") as f:
                            f.write(driver.page_source)
                        save_screenshot(driver, screenshot_dir)
                        return
                if not any_checked:
                    print("Download Housekeep: 無任何 Checkbox 被選中 (備用定位)，可能影響 Email 按鈕", flush=True)
                    with open("page_source.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                    save_screenshot(driver, screenshot_dir)
            else:
                print("Download Housekeep: 備用定位也未找到 checkbox，任務中止", flush=True)
                print(f"當前 URL: {driver.current_url}", flush=True)
                print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
                with open("page_source.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                save_screenshot(driver, screenshot_dir)
                return

        print("Download Housekeep: 點擊 Email 按鈕...", flush=True)
        try:
            email_button = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@title='Email'] | //*[contains(@class, 'MuiIconButton-root')][@title='Email'] | //button[descendant::path[@fill='#0080ff']]")))
            is_disabled = email_button.get_attribute("disabled")
            print(f"Download Housekeep: Email 按鈕狀態 - 禁用: {is_disabled}", flush=True)
            if is_disabled:
                print("Download Housekeep: Email 按鈕被禁用，嘗試選擇 Owner、Report Type 或點擊 Generate/Submit/Query...", flush=True)
                try:
                    # 嘗試選擇 Owner
                    owner_select = wait.until(EC.presence_of_element_located((By.XPATH, "//select[@id='owner'] | //select[contains(@name, 'owner')] | //select[contains(@id, 'company')] | //select")))
                    owner_select.click()
                    driver.execute_script("arguments[0].value = arguments[0].options[1].value;", owner_select)
                    print("Download Housekeep: Owner 選擇完成", flush=True)
                    time.sleep(5)
                except TimeoutException:
                    print("Download Housekeep: 未找到 Owner 下拉選單，跳過", flush=True)

                try:
                    # 嘗試選擇 Report Type
                    report_type_select = wait.until(EC.presence_of_element_located((By.XPATH, "//select[@id='reportType'] | //select[contains(@name, 'reportType')] | //select[contains(@id, 'status')]")))
                    report_type_select.click()
                    driver.execute_script("arguments[0].value = arguments[0].options[1].value;", report_type_select)
                    print("Download Housekeep: Report Type 選擇完成", flush=True)
                    time.sleep(5)
                except TimeoutException:
                    print("Download Housekeep: 未找到 Report Type 下拉選單，跳過", flush=True)

                try:
                    # 嘗試點擊 Generate、Submit 或 Query 按鈕
                    action_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Generate')] | //button[contains(text(), 'Submit')] | //button[contains(text(), 'Query')] | //button[@aria-label='Generate'] | //button[@aria-label='Submit'] | //button[@aria-label='Query']")))
                    driver.execute_script("arguments[0].click();", action_button)
                    print("Download Housekeep: Generate/Submit/Query 按鈕點擊成功", flush=True)
                    time.sleep(5)
                except TimeoutException:
                    print("Download Housekeep: 未找到 Generate、Submit 或 Query 按鈕，跳過", flush=True)

                # 再次檢查 Email 按鈕
                email_button = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@title='Email'] | //*[contains(@class, 'MuiIconButton-root')][@title='Email'] | //button[descendant::path[@fill='#0080ff']]")))
                is_disabled = email_button.get_attribute("disabled")
                print(f"Download Housekeep: Email 按鈕狀態 (再次檢查) - 禁用: {is_disabled}", flush=True)
                if is_disabled:
                    print("Download Housekeep: Email 按鈕仍被禁用，任務中止", flush=True)
                    with open("page_source.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                    save_screenshot(driver, screenshot_dir)
                    return

            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Email'] | //*[contains(@class, 'MuiIconButton-root')][@title='Email'] | //button[descendant::path[@fill='#0080ff']]")))
            driver.execute_script("arguments[0].scrollIntoView(true);", email_button)
            email_button.click()
            print("Download Housekeep: Email 按鈕點擊成功", flush=True)

            target_email = os.environ.get('TARGET_EMAIL', 'paklun@ckline.com.hk')
            print("Download Housekeep: 輸入目標 Email 地址...", flush=True)
            email_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='to'] | //input[@id='to']")))
            email_field.clear()
            email_field.send_keys(target_email)
            print("Download Housekeep: Email 地址輸入完成", flush=True)

            current_time = datetime.now().strftime("%m:%d %H:%M")
            print(f"Download Housekeep: 輸入內文，格式為 {current_time}", flush=True)
            body_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='body'] | //textarea[@id='body']")))
            body_field.clear()
            body_field.send_keys(current_time)
            print("Download Housekeep: 內文輸入完成", flush=True)

            print("Download Housekeep: 點擊 Confirm 按鈕...", flush=True)
            confirm_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='EmailDialog']//button[contains(text(), 'Confirm')] | //*[@id='EmailDialog']//button[@type='submit']")))
            driver.execute_script("arguments[0].click();", confirm_button)
            print("Download Housekeep: Confirm 按鈕點擊成功", flush=True)
        except TimeoutException as e:
            print(f"Download Housekeep: Email 處理失敗: {str(e)}", flush=True)
            print(f"當前 URL: {driver.current_url}", flush=True)
            print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
            with open("page_source.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            save_screenshot(driver, screenshot_dir)
            return
    except Exception as e:
        print(f"Download Housekeep: 錯誤: {str(e)}", flush=True)
        print(f"當前 URL: {driver.current_url}", flush=True)
        print(f"頁面 HTML (前500字符): {driver.page_source[:500]}", flush=True)
        with open("page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        save_screenshot(driver, screenshot_dir)
        return

def main():
    # 創建截圖目錄
    screenshot_dir = "./screenshots"
    if not os.path.exists(screenshot_dir):
        os.makedirs(screenshot_dir)
        print(f"已創建截圖目錄: {screenshot_dir}", flush=True)

    setup_environment()
    chromedriver_path = check_chromium_compatibility()
    driver = None
    try:
        driver = webdriver.Chrome(service=Service(chromedriver_path), options=get_chrome_options())
        driver.set_page_load_timeout(90)
        print("CPLUS WebDriver 初始化成功", flush=True)

        company_code = os.environ.get('COMPANY_CODE', 'CKL')
        user_id = os.environ.get('USER_ID', 'KEN')
        password = os.environ.get('SITE_PASSWORD')
        if not password:
            raise ValueError("環境變量 SITE_PASSWORD 未設置")

        login_cplus(driver, company_code, user_id, password, screenshot_dir)
        download_housekeep_report(driver, screenshot_dir)

    except Exception as e:
        print(f"Download Housekeep 錯誤: {str(e)}", flush=True)
        if driver:
            save_screenshot(driver, screenshot_dir)
    finally:
        if driver:
            try:
                print("Download Housekeep: 嘗試登出...", flush=True)
                try:
                    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
                    print("Download Housekeep: 檢測到登出按鈕，假設已登錄", flush=True)
                except TimeoutException:
                    print("Download Housekeep: 未檢測到登出按鈕，假設未登錄，跳過登出", flush=True)
                    driver.quit()
                    print("Download Housekeep WebDriver 關閉", flush=True)
                    return

                logout_menu_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
                driver.execute_script("arguments[0].click();", logout_menu_button)
                print("Download Housekeep: 登錄按鈕點擊成功", flush=True)

                try:
                    logout_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='menu-list-grow']/div[6]/li")))
                    driver.execute_script("arguments[0].click();", logout_option)
                    print("Download Housekeep: Logout 選項點擊成功", flush=True)

                    time.sleep(30)
                    final_url = driver.current_url
                    print(f"Download Housekeep: 登出後 URL: {final_url}", flush=True)
                except TimeoutException:
                    print("Download Housekeep: 登出選項未找到，跳過登出", flush=True)
            except Exception as logout_error:
                print(f"Download Housekeep: 登出失敗: {str(logout_error)}", flush=True)
            driver.quit()
            print("Download Housekeep WebDriver 關閉", flush=True)

        # 壓縮截圖為 ZIP 並標記為 artifact
        if os.path.exists(screenshot_dir):
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_path = f"screenshots_{timestamp}.zip"
            shutil.make_archive(zip_path[:-4], 'zip', screenshot_dir)
            print(f"已創建截圖 ZIP 文件: {zip_path}", flush=True)
            # 在 GitHub Actions 中，ZIP 文件將作為 artifact 上傳
            # 假設使用 actions/upload-artifact，需在 workflow 文件中配置：
            # - uses: actions/upload-artifact@v3
            #   with:
            #     name: screenshots
            #     path: screenshots_*.zip

if __name__ == "__main__":
    main()
    print("Download Housekeep 腳本完成", flush=True)
