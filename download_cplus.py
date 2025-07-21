import os
import time
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
from selenium.common.exceptions import TimeoutException

# 設置下載目錄
download_dir = os.path.abspath("downloads")
if not os.path.exists(download_dir):
    os.makedirs(download_dir)
    print(f"創建下載目錄: {download_dir}", flush=True)

# 確保環境準備
def setup_environment():
    try:
        subprocess.run(['sudo', 'apt-get', 'update', '-qq'], check=True)
        subprocess.run(['sudo', 'apt-get', 'install', '-y', 'chromium-browser', 'chromium-chromedriver'], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.run(['pip', 'install', '--upgrade', 'pip'], check=True)
        subprocess.run(['pip', 'install', 'selenium'], check=True)
        print("環境準備完成", flush=True)
    except subprocess.CalledProcessError as e:
        print(f"環境準備失敗: {e}", flush=True)
        raise

# 設置 Chrome 選項
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--disable-popup-blocking')
chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
prefs = {"download.default_directory": download_dir, "download.prompt_for_download": False, "safebrowsing.enabled": False}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.binary_location = '/usr/bin/chromium-browser'

# 初始化 WebDriver
print("嘗試初始化 WebDriver...", flush=True)
try:
    setup_environment()
    driver = webdriver.Chrome(options=chrome_options)
    print("WebDriver 初始化成功", flush=True)
except Exception as e:
    print(f"WebDriver 初始化失敗: {str(e)}", flush=True)
    raise

try:
    # 前往登入頁面 (CPLUS)
    print("嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/", flush=True)
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    print(f"網站已成功打開，當前 URL: {driver.current_url}", flush=True)
    time.sleep(2)

    # 點擊登錄前嘅按鈕 (CPLUS)
    print("點擊登錄前按鈕...", flush=True)
    wait = WebDriverWait(driver, 20)
    login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
    login_button_pre.click()
    print("登錄前按鈕點擊成功", flush=True)
    time.sleep(2)

    # 輸入 COMPANY CODE (CPLUS)
    print("輸入 COMPANY CODE...", flush=True)
    company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    company_code_field.send_keys("CKL")
    print("COMPANY CODE 輸入完成", flush=True)
    time.sleep(1)

    # 輸入 USER ID (CPLUS)
    print("輸入 USER ID...", flush=True)
    user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
    user_id_field.send_keys("KEN")
    print("USER ID 輸入完成", flush=True)
    time.sleep(1)

    # 輸入 PASSWORD (CPLUS)
    print("輸入 PASSWORD...", flush=True)
    password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
    password_field.send_keys(os.environ.get('SITE_PASSWORD'))
    print("PASSWORD 輸入完成", flush=True)
    time.sleep(1)

    # 點擊 LOGIN 按鈕 (CPLUS)
    print("點擊 LOGIN 按鈕...", flush=True)
    login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
    login_button.click()
    print("LOGIN 按鈕點擊成功", flush=True)
    time.sleep(2)

    # 前往 Container Movement Log 頁面 (CPLUS)
    print("直接前往 Container Movement Log...", flush=True)
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
    time.sleep(2)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    print("Container Movement Log 頁面加載完成", flush=True)
    time.sleep(2)

    # 點擊 Search (CPLUS)
    print("點擊 Search...", flush=True)
    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button/span[1]")))
    search_button.click()
    print("Search 按鈕點擊成功", flush=True)
    time.sleep(5)

    # 點擊 Download (CPLUS)
    print("點擊 Download...", flush=True)
    download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button")))
    download_button.click()
    print("Download 按鈕點擊成功", flush=True)
    time.sleep(90)  # 延長下載等待時間

    # 檢查 Container Movement Log 下載文件
    print("檢查 Container Movement Log 下載文件...", flush=True)
    start_time = time.time()
    downloaded_files = []
    while time.time() - start_time < 120:  # 延長檢查時間
        downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
        if downloaded_files:
            break
        time.sleep(2)
    if downloaded_files:
        print(f"Container Movement Log 下載完成，檔案位於: {download_dir}", flush=True)
        for file in downloaded_files:
            print(f"找到檔案: {file}", flush=True)

    # 前往 OnHandContainerList 頁面 (CPLUS)
    print("前往 OnHandContainerList 頁面...", flush=True)
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/OnHandContainerList")
    time.sleep(2)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    print("OnHandContainerList 頁面加載完成", flush=True)
    time.sleep(2)

    # 點擊 Search (CPLUS)
    print("點擊 Search...", flush=True)
    search_button_onhand = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")))
    search_button_onhand.click()
    print("Search 按鈕點擊成功", flush=True)
    time.sleep(5)

    # 點擊 Export (CPLUS)
    print("點擊 Export...", flush=True)
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button")))
    export_button.click()
    print("Export 按鈕點擊成功", flush=True)
    time.sleep(2)

    # 點擊 Export as CSV (CPLUS)
    print("點擊 Export as CSV...", flush=True)
    export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']")))
    export_csv_button.click()
    print("Export as CSV 按鈕點擊成功", flush=True)
    time.sleep(90)  # 延長下載等待時間

    # 檢查 OnHandContainerList 下載文件
    print("檢查 OnHandContainerList 下載文件...", flush=True)
    start_time = time.time()
    while time.time() - start_time < 120:  # 延長檢查時間
        downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
        if downloaded_files:
            break
        time.sleep(5)
    if downloaded_files:
        print(f"OnHandContainerList 下載完成，檔案位於: {download_dir}", flush=True)
        for file in downloaded_files:
            print(f"找到檔案: {file}", flush=True)

    # 前往 barge.oneport.com 登入頁面
    print("嘗試打開網站 https://barge.oneport.com/bargeBooking...", flush=True)
    driver.get("https://barge.oneport.com/bargeBooking")
    print(f"網站已成功打開，當前 URL: {driver.current_url}", flush=True)
    time.sleep(3)

    # 輸入 COMPANY ID
    print("輸入 COMPANY ID...", flush=True)
    company_id_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='mat-input-0']")))
    company_id_field.send_keys("CKL")
    print("COMPANY ID 輸入完成", flush=True)
    time.sleep(1)

    # 輸入 USER ID
    print("輸入 USER ID...", flush=True)
    user_id_field_barge = driver.find_element(By.XPATH, "//*[@id='mat-input-1']")
    user_id_field_barge.send_keys("barge")
    print("USER ID 輸入完成", flush=True)
    time.sleep(1)

    # 輸入 PW
    print("輸入 PW...", flush=True)
    password_field_barge = driver.find_element(By.XPATH, "//*[@id='mat-input-2']")
    password_field_barge.send_keys("123456")
    print("PW 輸入完成", flush=True)
    time.sleep(1)

    # 點擊 LOGIN
    print("點擊 LOGIN 按鈕...", flush=True)
    login_button_barge = driver.find_element(By.XPATH, "//*[@id='login-form-container']/app-login-form/form/div/button")
    login_button_barge.click()
    print("LOGIN 按鈕點擊成功", flush=True)
    time.sleep(3)

    # 點擊主工具欄
    print("點擊主工具欄...", flush=True)
    toolbar_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[1]/span[1]/mat-icon")))
    toolbar_button.click()
    print("主工具欄點擊成功", flush=True)
    time.sleep(2)

    # 點擊 Report
    print("點擊 Report...", flush=True)
    report_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-menu-panel-4']/div/button[4]/span")))
    report_button.click()
    print("Report 點擊成功", flush=True)
    time.sleep(2)

    # 選擇 Report Type
    print("選擇 Report Type...", flush=True)
    report_type_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-select-value-61']/span")))
    report_type_select.click()
    print("Report Type 選擇開始", flush=True)
    time.sleep(2)

    # 點擊 Container Detail
    print("點擊 Container Detail...", flush=True)
    container_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-option-508']/span")))
    container_detail_option.click()
    print("Container Detail 點擊成功", flush=True)
    time.sleep(5)

    # 點擊 Download
    print("點擊 Download...", flush=True)
    download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='content-mount']/app-download-report/div[2]/div/form/div[2]/button")))
    download_button_barge.click()
    print("Download 按鈕點擊成功", flush=True)
    time.sleep(90)  # 延長下載等待時間

    # 檢查 Barge Container Detail 下載文件
    print("檢查 Barge Container Detail 下載文件...", flush=True)
    start_time = time.time()
    while time.time() - start_time < 120:  # 延長檢查時間
        downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
        if downloaded_files:
            break
        time.sleep(5)
    if downloaded_files:
        print(f"Barge Container Detail 下載完成，檔案位於: {download_dir}", flush=True)
        for file in downloaded_files:
            print(f"找到檔案: {file}", flush=True)

    # 點擊工具欄進行登出
    print("點擊工具欄進行登出...", flush=True)
    logout_toolbar = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
    driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar)  # 滾動到元素
    time.sleep(1)
    driver.execute_script("arguments[0].click();", logout_toolbar)  # 使用 JavaScript 點擊
    print("工具欄點擊成功", flush=True)
    time.sleep(2)
    
    # 點擊 Logout
    print("點擊 Logout...", flush=True)
    logout_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-menu-panel-11']/div/button/span")))
    driver.execute_script("arguments[0].scrollIntoView(true);", logout_button)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", logout_button)
    print("Logout 點擊成功", flush=True)
    time.sleep(5)

    # 檢查所有下載文件
    print("檢查所有下載文件...", flush=True)
    start_time = time.time()
    while time.time() - start_time < 120:  # 延長檢查時間
        downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
        if downloaded_files:
            break
        time.sleep(5)
    if downloaded_files:
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

            # 創建郵件
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = receiver_email
            msg['Subject'] = f"[TESTING]HIT DAILY + {datetime.now().strftime('%Y-%m-%d')}"

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

except Exception as e:
    print(f"發生錯誤: {str(e)}", flush=True)
    raise

finally:
    driver.quit()
