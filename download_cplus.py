import os
import time
import subprocess
import threading
import traceback
from datetime import datetime
import pytz
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
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# 設定香港時區
hkt = pytz.timezone('Asia/Hong_Kong')

# 全局變量
download_dir = os.path.abspath("downloads")
if not os.path.exists(download_dir):
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
        if "selenium" not in result.stdout or "webdriver-manager" not in subprocess.run(['pip', 'show', 'webdriver-manager'], capture_output=True, text=True).stdout or "pytz" not in subprocess.run(['pip', 'show', 'pytz'], capture_output=True, text=True).stdout:
            subprocess.run(['pip', 'install', 'selenium', 'webdriver-manager', 'pytz'], check=True)
            print("Selenium、WebDriver Manager 及 pytz 已安裝", flush=True)
        else:
            print("Selenium、WebDriver Manager 及 pytz 已存在，跳過安裝", flush=True)
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
    prefs = {"download.default_directory": download_dir, "download.prompt_for_download": False, "safebrowsing.enabled": False}
    chrome_options.add_experimental_option("prefs", prefs)
    binary_path = os.path.expanduser('~/chromium-bin/chromium-browser')
    if os.path.exists(binary_path):
        chrome_options.binary_location = binary_path
    else:
        print(f"警告: Chrome 二進制文件 {binary_path} 不存在，使用系統預設", flush=True)
    return chrome_options

# CPLUS 操作
def process_cplus():
    driver = None
    try:
        driver = webdriver.Chrome(options=get_chrome_options())
        print("CPLUS WebDriver 初始化成功", flush=True)

        # 前往登入頁面 (CPLUS)
        print("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/", flush=True)
        driver.get("https://cplus.hit.com.hk/frontpage/#/")
        print(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)
        time.sleep(2)

        # 點擊登錄前嘅按鈕 (CPLUS)
        print("CPLUS: 點擊登錄前按鈕...", flush=True)
        wait = WebDriverWait(driver, 20)
        login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
        login_button_pre.click()
        print("CPLUS: 登錄前按鈕點擊成功", flush=True)
        time.sleep(2)

        # 輸入 COMPANY CODE (CPLUS)
        print("CPLUS: 輸入 COMPANY CODE...", flush=True)
        company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
        company_code_field.send_keys("CKL")
        print("CPLUS: COMPANY CODE 輸入完成", flush=True)
        time.sleep(1)

        # 輸入 USER ID (CPLUS)
        print("CPLUS: 輸入 USER ID...", flush=True)
        user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
        user_id_field.send_keys("KEN")
        print("CPLUS: USER ID 輸入完成", flush=True)
        time.sleep(1)

        # 輸入 PASSWORD (CPLUS)
        print("CPLUS: 輸入 PASSWORD...", flush=True)
        password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
        password_field.send_keys(os.environ.get('SITE_PASSWORD'))
        print("CPLUS: PASSWORD 輸入完成", flush=True)
        time.sleep(1)

        # 點擊 LOGIN 按鈕 (CPLUS)
        print("CPLUS: 點擊 LOGIN 按鈕...", flush=True)
        login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
        login_button.click()
        print("CPLUS: LOGIN 按鈕點擊成功", flush=True)
        time.sleep(2)

        # 前往 Container Movement Log 頁面 (CPLUS)
        print("CPLUS: 直接前往 Container Movement Log...", flush=True)
        driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
        time.sleep(3)  # 增加等待時間至 3 秒
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]//form")))
            print("CPLUS: Container Movement Log 頁面加載完成", flush=True)
        except TimeoutException as e:
            print(f"CPLUS: 頁面加載失敗: {str(e)}", flush=True)
            raise

        # 點擊 Search (CPLUS)
        print("CPLUS: 點擊 Search...", flush=True)
        try:
            search_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button")))
            search_button.click()
            print("CPLUS: Search 按鈕點擊成功", flush=True)
        except TimeoutException:
            print("CPLUS: Search 按鈕未找到，嘗試備用定位 1...", flush=True)
            try:
                search_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and .//span[contains(text(), 'Search')]]")))
                search_button.click()
                print("CPLUS: 備用 Search 按鈕 1 點擊成功", flush=True)
            except TimeoutException:
                print("CPLUS: 備用 Search 按鈕 1 失敗，嘗試備用定位 2...", flush=True)
                search_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]")))
                search_button.click()
                print("CPLUS: 備用 Search 按鈕 2 點擊成功", flush=True)
        time.sleep(5)

        # 點擊 Download (CPLUS)
        print("CPLUS: 點擊 Download...", flush=True)
        download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button")))
        download_button.click()
        print("CPLUS: Download 按鈕點擊成功", flush=True)

        # 等待下載完成 (假設有成功提示或按鈕禁用)
        try:
            WebDriverWait(driver, 60).until(
                EC.invisibility_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button"))
                or EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Download Complete')]"))
            )
            print("CPLUS: Container Movement Log 下載完成", flush=True)
        except TimeoutException:
            print("CPLUS: 下載完成提示未找到，繼續檢查文件...", flush=True)

        # 檢查文件
        start_time = time.time()
        while time.time() - start_time < 30:
            downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
            if downloaded_files:
                print(f"CPLUS: Container Movement Log 下載完成，檔案位於: {download_dir}", flush=True)
                for file in downloaded_files:
                    print(f"CPLUS: 找到檔案: {file}", flush=True)
                break
            time.sleep(2)

        # 前往 OnHandContainerList 頁面 (CPLUS)
        print("CPLUS: 前往 OnHandContainerList 頁面...", flush=True)
        driver.get("https://cplus.hit.com.hk/app/#/enquiry/OnHandContainerList")
        time.sleep(2)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
        print("CPLUS: OnHandContainerList 頁面加載完成", flush=True)
        time.sleep(2)

        # 點擊 Search (CPLUS)
        print("CPLUS: 點擊 Search...", flush=True)
        try:
            search_button_onhand = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")))
            search_button_onhand.click()
            print("CPLUS: Search 按鈕點擊成功", flush=True)
        except TimeoutException:
            print("CPLUS: Search 按鈕未找到，嘗試備用定位...", flush=True)
            search_button_onhand = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]")))
            search_button_onhand.click()
            print("CPLUS: 備用 Search 按鈕點擊成功", flush=True)
        time.sleep(5)

        # 點擊 Export (CPLUS)
        print("CPLUS: 點擊 Export...", flush=True)
        export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button")))
        export_button.click()
        print("CPLUS: Export 按鈕點擊成功", flush=True)
        time.sleep(2)

        # 點擊 Export as CSV (CPLUS)
        print("CPLUS: 點擊 Export as CSV...", flush=True)
        export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']")))
        export_csv_button.click()
        print("CPLUS: Export as CSV 按鈕點擊成功", flush=True)

        # 等待下載完成 (假設有成功提示或按鈕禁用)
        try:
            WebDriverWait(driver, 60).until(
                EC.invisibility_of_element_located((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']"))
                or EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Export Complete')]"))
            )
            print("CPLUS: OnHandContainerList 下載完成", flush=True)
        except TimeoutException:
            print("CPLUS: 下載完成提示未找到，繼續檢查文件...", flush=True)

        # 檢查文件
        start_time = time.time()
        while time.time() - start_time < 30:
            downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
            if downloaded_files:
                print(f"CPLUS: OnHandContainerList 下載完成，檔案位於: {download_dir}", flush=True)
                for file in downloaded_files:
                    print(f"CPLUS: 找到檔案: {file}", flush=True)
                break
            time.sleep(2)

    except Exception as e:
        print(f"CPLUS 錯誤: {str(e)} - 堆棧跟踪: {traceback.format_exc()}", flush=True)

    finally:
        # 確保登出
        try:
            if driver:
                print("CPLUS: 嘗試登出...", flush=True)
                logout_menu_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_menu_button)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", logout_menu_button)
                print("CPLUS: 登錄按鈕點擊成功", flush=True)

                logout_option = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='menu-list-grow']/div[6]/li")))
                driver.execute_script("arguments[0].scrollIntoView(true);", logout_option)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", logout_option)
                print("CPLUS: Logout 選項點擊成功", flush=True)
                time.sleep(5)
        except Exception as logout_error:
            print(f"CPLUS: 登出失敗: {str(logout_error)} - 堆棧跟踪: {traceback.format_exc()}", flush=True)

        if driver:
            driver.quit()
            print("CPLUS WebDriver 關閉", flush=True)

# Barge 操作
def process_barge():
    driver = None
    try:
        driver = webdriver.Chrome(options=get_chrome_options())
        print("Barge WebDriver 初始化成功", flush=True)

        # 前往登入頁面 (Barge)
        print("Barge: 嘗試打開網站 https://barge.oneport.com/bargeBooking...", flush=True)
        driver.get("https://barge.oneport.com/bargeBooking")
        print(f"Barge: 網站已成功打開，當前 URL: {driver.current_url}", flush=True)
        time.sleep(3)

        # 輸入 COMPANY ID
        print("Barge: 輸入 COMPANY ID...", flush=True)
        wait = WebDriverWait(driver, 20)
        company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='mat-input-0']")))
        company_id_field.send_keys("CKL")
        print("Barge: COMPANY ID 輸入完成", flush=True)
        time.sleep(1)

        # 輸入 USER ID
        print("Barge: 輸入 USER ID...", flush=True)
        user_id_field_barge = driver.find_element(By.XPATH, "//*[@id='mat-input-1']")
        user_id_field_barge.send_keys("barge")
        print("Barge: USER ID 輸入完成", flush=True)
        time.sleep(1)

        # 輸入 PW
        print("Barge: 輸入 PW...", flush=True)
        password_field_barge = driver.find_element(By.XPATH, "//*[@id='mat-input-2']")
        password_field_barge.send_keys("123456")
        print("Barge: PW 輸入完成", flush=True)
        time.sleep(1)

        # 點擊 LOGIN
        print("Barge: 點擊 LOGIN 按鈕...", flush=True)
        login_button_barge = driver.find_element(By.XPATH, "//*[@id='login-form-container']/app-login-form/form/div/button")
        login_button_barge.click()
        print("Barge: LOGIN 按鈕點擊成功", flush=True)
        time.sleep(3)

        # 點擊主工具欄
        print("Barge: 點擊主工具欄...", flush=True)
        toolbar_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[1]/span[1]/mat-icon")))
        toolbar_button.click()
        print("Barge: 主工具欄點擊成功", flush=True)
        time.sleep(8)  # 增加等待時間至 8 秒
        # 等待菜單出現並包含 Report 按鈕
        WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.XPATH, "//div[@role='menu']//button")))
        WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.XPATH, "//button[descendant::span[text()='Reports']]")))

        # 點擊 Report
        print("Barge: 點擊 Report...", flush=True)
        try:
            report_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[descendant::span[text()='Reports']]")))
            report_button.click()
            print("Barge: Report 點擊成功", flush=True)
        except TimeoutException:
            print("Barge: Report 按鈕未找到，嘗試備用定位...", flush=True)
            try:
                report_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[./span[contains(text(), 'Reports')]]")))
                report_button.click()
                print("Barge: 備用 Report 按鈕點擊成功", flush=True)
            except TimeoutException:
                print("Barge: 備用 Report 按鈕失敗，跳過...", flush=True)
                raise
        time.sleep(2)

        # 選擇 Report Type
        print("Barge: 選擇 Report Type...", flush=True)
        report_type_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-select-value-61']/span")))
        report_type_select.click()
        print("Barge: Report Type 選擇開始", flush=True)
        time.sleep(2)

        # 點擊 Container Detail
        print("Barge: 點擊 Container Detail...", flush=True)
        container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-option-508']/span")))
        container_detail_option.click()
        print("Barge: Container Detail 點擊成功", flush=True)
        time.sleep(5)

        # 點擊 Download
        print("Barge: 點擊 Download...", flush=True)
        download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='content-mount']/app-download-report/div[2]/div/form/div[2]/button")))
        download_button_barge.click()
        print("Barge: Download 按鈕點擊成功", flush=True)

        # 等待下載完成 (假設有成功提示或按鈕禁用)
        try:
            WebDriverWait(driver, 60).until(
                EC.invisibility_of_element_located((By.XPATH, "//*[@id='content-mount']/app-download-report/div[2]/div/form/div[2]/button"))
                or EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Download Complete')]"))
            )
            print("Barge: Container Detail 下載完成", flush=True)
        except TimeoutException:
            print("Barge: 下載完成提示未找到，繼續檢查文件...", flush=True)

        # 檢查文件
        start_time = time.time()
        while time.time() - start_time < 30:
            downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
            if downloaded_files:
                print(f"Barge: Container Detail 下載完成，檔案位於: {download_dir}", flush=True)
                for file in downloaded_files:
                    print(f"Barge: 找到檔案: {file}", flush=True)
                break
            time.sleep(2)

        # 登出 Barge
        print("Barge: 點擊工具欄進行登出...", flush=True)
        try:
            logout_toolbar_barge = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
            driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar_barge)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", logout_toolbar_barge)
            print("Barge: 工具欄點擊成功", flush=True)
        except TimeoutException:
            print("Barge: 主工具欄登出按鈕未找到，嘗試備用定位...", flush=True)
            raise

        print("Barge: 點擊 Logout 選項...", flush=True)
        try:
            logout_button_barge = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-menu-panel-11']/div/button/span")))
            driver.execute_script("arguments[0].scrollIntoView(true);", logout_button_barge)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", logout_button_barge)
            print("Barge: Logout 選項點擊成功", flush=True)
        except TimeoutException:
            print("Barge: Logout 選項未找到，嘗試備用定位...", flush=True)
            raise

        time.sleep(5)

    except Exception as e:
        print(f"Barge 錯誤: {str(e)} - 堆棧跟踪: {traceback.format_exc()}", flush=True)

    finally:
        if driver:
            driver.quit()
            print("Barge WebDriver 關閉", flush=True)

# 主函數
if __name__ == "__main__":
    # 設定系統時區為 HKT
    os.environ['TZ'] = 'Asia/Hong_Kong'
    time.tzset()

    # 啟動兩個線程
    cplus_thread = threading.Thread(target=process_cplus)
    barge_thread = threading.Thread(target=process_barge)

    cplus_thread.start()
    barge_thread.start()

    # 等待兩個線程完成
    cplus_thread.join()
    barge_thread.join()

    # 檢查所有下載文件
    print("檢查所有下載文件...", flush=True)
    start_time = time.time()
    while time.time() - start_time < 120:
        downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
        if downloaded_files:
            break
        time.sleep(5)
    if downloaded_files:
        print(f"所有下載完成，檔案位於: {download_dir}", flush=True)
        for file in downloaded_files:
            print(f"找到檔案: {file}", flush=True)

        # 發送 Zoho Mail (使用 HKT 時間)
        hkt_time = datetime.now(hkt).strftime('%Y-%m-%d %H:%M:%S')
        print(f"郵件發送時間 (HKT): {hkt_time}", flush=True)
        try:
            smtp_server = 'smtp.zoho.com'
            smtp_port = 587
            sender_email = os.environ.get('ZOHO_EMAIL', 'paklun_ckline@zohomail.com')
            sender_password = os.environ.get('ZOHO_PASSWORD', '@d6G.Pie5UkEPqm')
            receiver_email = 'paklun@ckline.com.hk'  # 修復為正確的郵件地址

            # 創建郵件
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = receiver_email
            msg['Subject'] = f"[TESTING]HIT DAILY {datetime.now(hkt).strftime('%Y-%m-%d')}"

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
            print("郵件發送成功!", flush=True)
        except Exception as e:
            print(f"郵件發送失敗: {str(e)}", flush=True)
    else:
        print("所有下載失敗，無文件可發送", flush=True)

    print("腳本完成", flush=True)
