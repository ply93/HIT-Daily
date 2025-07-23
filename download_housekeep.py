import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

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
    chrome_options.binary_location = '/usr/bin/chromium-browser'
    return chrome_options

# 主任務邏輯
def process_download_housekeep():
    driver = None
    try:
        driver = webdriver.Chrome(options=get_chrome_options())
        print("Download Housekeep WebDriver 初始化成功", flush=True)

        # 前往登錄頁面
        print("Download Housekeep: 嘗試打開網站 https://cplus.hit.com.hk/app/#/", flush=True)
        driver.get("https://cplus.hit.com.hk/app/#/")
        print(f"Download Housekeep: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)
        time.sleep(2)

        # 點擊登錄前按鈕
        print("Download Housekeep: 點擊登錄前按鈕...", flush=True)
        wait = WebDriverWait(driver, 20)
        login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
        login_button_pre.click()
        print("Download Housekeep: 登錄前按鈕點擊成功", flush=True)
        time.sleep(2)

        # 輸入 COMPANY CODE
        print("Download Housekeep: 輸入 COMPANY CODE...", flush=True)
        company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
        company_code_field.send_keys("CKL")
        print("Download Housekeep: COMPANY CODE 輸入完成", flush=True)
        time.sleep(1)

        # 輸入 USER ID
        print("Download Housekeep: 輸入 USER ID...", flush=True)
        user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
        user_id_field.send_keys("KEN")
        print("Download Housekeep: USER ID 輸入完成", flush=True)
        time.sleep(1)

        # 輸入 PASSWORD
        print("Download Housekeep: 輸入 PASSWORD...", flush=True)
        password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
        password_field.send_keys(os.environ.get('SITE_PASSWORD'))
        print("Download Housekeep: PASSWORD 輸入完成", flush=True)
        time.sleep(1)

        # 點擊 LOGIN 按鈕
        print("Download Housekeep: 點擊 LOGIN 按鈕...", flush=True)
        login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
        login_button.click()
        print("Download Housekeep: LOGIN 按鈕點擊成功", flush=True)
        time.sleep(2)

        # 前往 Housekeep Report 頁面
        print("Download Housekeep: 前往 Housekeep Report 頁面...", flush=True)
        driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
        print("Download Housekeep: Housekeep Report 頁面加載完成", flush=True)
        time.sleep(2)

        # 檢查並設置日期
        today = datetime.now().strftime("%d/%m/%Y")  # 格式：23/07/2025
        print(f"Download Housekeep: 檢查日期，今日為 {today}", flush=True)
        
        # 檢查 <input id="from">
        try:
            from_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='from']")))
            from_value = from_field.get_attribute("value")
            if from_value != today:
                print(f"Download Housekeep: From 日期 ({from_value}) 不為今日，設置為 {today}", flush=True)
                from_field.clear()
                from_field.send_keys(today)
            else:
                print("Download Housekeep: From 日期已正確", flush=True)
        except TimeoutException:
            print("Download Housekeep: 未找到 From 輸入框，跳過日期檢查", flush=True)
        time.sleep(1)

        # 檢查 <input id="to">
        try:
            to_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='to']")))
            to_value = to_field.get_attribute("value")
            if to_value != today:
                print(f"Download Housekeep: To 日期 ({to_value}) 不為今日，設置為 {today}", flush=True)
                to_field.clear()
                to_field.send_keys(today)
            else:
                print("Download Housekeep: To 日期已正確", flush=True)
        except TimeoutException:
            print("Download Housekeep: 未找到 To 輸入框，跳過日期檢查", flush=True)
        time.sleep(1)

        # 點擊 Search 按鈕
        print("Download Housekeep: 嘗試點擊 Search 按鈕...", flush=True)
        try:
            search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Search']/following-sibling::div//button")))
            driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", search_button)
            print("Download Housekeep: Search 按鈕點擊成功", flush=True)
            time.sleep(5)  # 等待報告加載
        except TimeoutException:
            print("Download Housekeep: Search 按鈕未找到，假設無需點擊", flush=True)

        # 動態查找並點擊所有 Email Excel checkbox
        print("Download Housekeep: 查找並點擊所有 Email Excel checkbox...", flush=True)
        try:
            checkboxes = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr/td[6]//input[@type='checkbox']")))
            print(f"Download Housekeep: 找到 {len(checkboxes)} 個 Email Excel checkbox", flush=True)
            
            if len(checkboxes) < 3 or len(checkboxes) > 6:
                print(f"Download Housekeep: Checkbox 數量 {len(checkboxes)} 不在預期範圍 (3-6)，請檢查", flush=True)

            for index, checkbox in enumerate(checkboxes, 1):
                try:
                    if checkbox.is_enabled() and checkbox.is_displayed() and not checkbox.is_selected():
                        driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                        time.sleep(0.5)
                        driver.execute_script("arguments[0].click();", checkbox)
                        print(f"Download Housekeep: Checkbox {index} 點擊成功", flush=True)
                    else:
                        print(f"Download Housekeep: Checkbox {index} 已選中或不可點擊，跳過", flush=True)
                except Exception as e:
                    print(f"Download Housekeep: Checkbox {index} 點擊失敗: {str(e)}", flush=True)
            time.sleep(2)
        except TimeoutException:
            print("Download Housekeep: 未找到 Email Excel checkbox，嘗試備用定位...", flush=True)
            checkboxes = driver.find_elements(By.CSS_SELECTOR, "input.jss140")
            if checkboxes:
                print(f"Download Housekeep: 備用定位找到 {len(checkboxes)} 個 checkbox", flush=True)
                for index, checkbox in enumerate(checkboxes, 1):
                    try:
                        if checkbox.is_enabled() and checkbox.is_displayed() and not checkbox.is_selected():
                            driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                            time.sleep(0.5)
                            driver.execute_script("arguments[0].click();", checkbox)
                            print(f"Download Housekeep: Checkbox {index} 點擊成功 (備用定位)", flush=True)
                    except Exception as e:
                        print(f"Download Housekeep: Checkbox {index} 點擊失敗 (備用定位): {str(e)}", flush=True)
            else:
                print("Download Housekeep: 備用定位也未找到 checkbox，任務中止", flush=True)
                return

        # 點擊 Email 按鈕
        print("Download Housekeep: 點擊 Email 按鈕...", flush=True)
        try:
            email_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/form/div[1]/div[8]/div/div/div[1]/div[1]/div[1]/button")))
            driver.execute_script("arguments[0].scrollIntoView(true);", email_button)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", email_button)
            print("Download Housekeep: Email 按鈕點擊成功", flush=True)
            time.sleep(2)

            # 輸入 Email 地址
            print("Download Housekeep: 輸入目標 Email 地址...", flush=True)
            email_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='to']")))
            email_field.clear()  # 清空現有內容
            email_field.send_keys("paklun@ckline.com.hk")
            print("Download Housekeep: Email 地址輸入完成", flush=True)
            time.sleep(1)

            # 輸入內文（今日日期和時間，格式 MM:DD XX:XX）
            current_time = datetime.now().strftime("%m:%d %H:%M")  # 例如 07:23 16:39
            print(f"Download Housekeep: 輸入內文，格式為 {current_time}", flush=True)
            body_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='body']")))
            body_field.clear()  # 清空現有內容
            body_field.send_keys(current_time)
            print("Download Housekeep: 內文輸入完成", flush=True)
            time.sleep(1)

            # 點擊 Confirm 按鈕
            print("Download Housekeep: 點擊 Confirm 按鈕...", flush=True)
            confirm_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='EmailDialog']/div[3]/div/form/div[3]/button[1]")))
            driver.execute_script("arguments[0].click();", confirm_button)
            print("Download Housekeep: Confirm 按鈕點擊成功", flush=True)
            time.sleep(2)

            # 檢查發送成功提示
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'successfully')]")))
                print("Download Housekeep: Email 發送成功", flush=True)
            except TimeoutException:
                print("Download Housekeep: 未找到發送成功提示，但繼續執行", flush=True)
        except TimeoutException as e:
            print(f"Download Housekeep: Email 處理失敗: {str(e)}", flush=True)
            return
        except NoSuchElementException as e:
            print(f"Download Housekeep: Email 按鈕或對話框元素未找到: {str(e)}", flush=True)
            return

    except Exception as e:
        print(f"Download Housekeep 錯誤: {str(e)}", flush=True)

    finally:
        if driver:
            try:
                # 嘗試登出
                print("Download Housekeep: 嘗試登出...", flush=True)
                logout_menu_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_menu_button)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", logout_menu_button)
                print("Download Housekeep: 登錄按鈕點擊成功", flush=True)

                logout_option = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='menu-list-grow']/div[6]/li")))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_option)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", logout_option)
                print("Download Housekeep: Logout 選項點擊成功", flush=True)
                time.sleep(2)
            except Exception as logout_error:
                print(f"Download Housekeep: 登出失敗: {str(logout_error)}", flush=True)
            driver.quit()
            print("Download Housekeep WebDriver 關閉", flush=True)

if __name__ == "__main__":
    process_download_housekeep()
    print("Download Housekeep 腳本完成", flush=True)
