import os
import time
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

# 設置下載目錄
download_dir = "downloads"
if not os.path.exists(download_dir):
    os.makedirs(download_dir)
    print(f"創建下載目錄: {download_dir}")

# 設置 Chrome 選項
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--remote-debugging-port=9222')
chrome_options.add_argument('--window-size=1920,1080')
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--allow-running-insecure-content')
chrome_options.add_argument('--disable-features=DownloadBubble')
chrome_options.add_argument('--disable-popup-blocking')
chrome_options.add_argument('--disable-save-password-bubble')
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "safebrowsing.enabled": True,
    "profile.default_content_settings.popups": 0,
    "directory_upgrade": True
}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.binary_location = '/usr/bin/chromium-browser'

# 初始化 WebDriver
print("嘗試初始化 WebDriver...")
try:
    driver = webdriver.Chrome(options=chrome_options)
    print("WebDriver 初始化成功")
except Exception as e:
    print(f"WebDriver 初始化失敗: {str(e)}")
    with open("error_log.txt", "a") as f:
        f.write(f"WebDriver 初始化失敗: {str(e)} - {datetime.now()}\n")
    raise

try:
    # 前往登入頁面
    print("嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/...")
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    print(f"網站已成功打開，當前 URL: {driver.current_url}")
    time.sleep(2)

    # 點擊登錄前嘅按鈕
    print("點擊登錄前按鈕...")
    wait = WebDriverWait(driver, 20)
    login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
    driver.execute_script("arguments[0].click();", login_button_pre)  # 用 JS 點擊
    print("登錄前按鈕點擊成功")
    time.sleep(2)

    # 輸入 COMPANY CODE
    print("輸入 COMPANY CODE...")
    company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    company_code_field.send_keys("CKL")
    print("COMPANY CODE 輸入完成")
    time.sleep(1)

    # 輸入 USER ID
    print("輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
    user_id_field.send_keys("KEN")
    print("USER ID 輸入完成")
    time.sleep(1)

    # 輸入 PASSWORD
    print("輸入 PASSWORD...")
    password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
    password_field.send_keys(os.environ.get('SITE_PASSWORD', 'default_password'))
    print("PASSWORD 輸入完成")
    time.sleep(1)

    # 點擊 LOGIN 按鈕
    print("點擊 LOGIN 按鈕...")
    login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
    driver.execute_script("arguments[0].click();", login_button)  # 用 JS 點擊
    print("LOGIN 按鈕點擊成功")
    time.sleep(10)

    # 檢查當前 URL 進行調試
    current_url = driver.current_url
    print(f"登錄後 URL: {current_url}")

    # 直接前往 Container Movement Log 頁面
    print("直接前往 Container Movement Log...")
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
    time.sleep(5)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    print("Container Movement Log 頁面加載完成")
    time.sleep(5)

    # 動態生成 expected_date (當日 23:59)
    expected_date = datetime.now().strftime("%d/%m/%Y 23:59")
    print("檢查 toDateTime...")
    to_date_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='toDateTime']")))
    to_date_value = to_date_field.get_attribute("value")
    print(f"toDateTime 值: {to_date_value}")
    if to_date_value == expected_date:
        print("toDateTime 匹配今日日期!")
    else:
        print(f"toDateTime 不匹配，預期 {expected_date}，實際 {to_date_value}")

    # 點擊 Search
    print("點擊 Search...")
    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button/span[1]")))
    driver.execute_script("arguments[0].click();", search_button)  # 用 JS 點擊
    print("Search 按鈕點擊成功")
    time.sleep(10)

    # 點擊 Download
    print("點擊 Download...")
    download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button")))
    driver.execute_script("arguments[0].click();", download_button)  # 用 JS 點擊
    print("Download 按鈕點擊成功")
    time.sleep(60)

    # 檢查下載文件（循環檢查）
    print("檢查下載文件...")
    start_time = time.time()
    downloaded_files = []
    while time.time() - start_time < 60:
        downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
        if downloaded_files:
            break
        time.sleep(5)
    if downloaded_files:
        print(f"下載完成，檔案位於: {download_dir}")
        for file in downloaded_files:
            print(f"找到檔案: {file}")
    else:
        print("下載失敗，無找到檔案")
        with open("error_log.txt", "a") as f:
            f.write(f"下載失敗，無找到檔案: {datetime.now()}\n")

    # 前往 OnHandContainerList 頁面
    print("前往 OnHandContainerList 頁面...")
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/OnHandContainerList")
    time.sleep(5)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    print("OnHandContainerList 頁面加載完成")
    time.sleep(5)

    # 點擊 Search
    print("點擊 Search...")
    search_button_onhand = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div[1]/form/div[1]/div[24]/div[2]/button/span[1]")))
    driver.execute_script("arguments[0].click();", search_button_onhand)  # 用 JS 點擊
    print("Search 按鈕點擊成功")
    time.sleep(10)

    # 點擊 Export
    print("點擊 Export...")
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div/span[1]/button")))
    driver.execute_script("arguments[0].click();", export_button)  # 用 JS 點擊
    print("Export 按鈕點擊成功")
    time.sleep(2)

    # 點擊 Export as CSV
    print("點擊 Export as CSV...")
    export_csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'MuiMenuItem-root') and text()='Export as CSV']")))
    driver.execute_script("arguments[0].click();", export_csv_button)  # 用 JS 點擊
    print("Export as CSV 按鈕點擊成功")
    time.sleep(60)

    # 檢查下載文件（循環檢查）
    print("檢查下載文件（包括 OnHandContainerList）...")
    start_time = time.time()
    while time.time() - start_time < 60:
        downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
        if downloaded_files:
            break
        time.sleep(5)
    if downloaded_files:
        print(f"下載完成，檔案位於: {download_dir}")
        for file in downloaded_files:
            print(f"找到檔案: {file}")
    else:
        print("下載失敗，無找到檔案")
        with open("error_log.txt", "a") as f:
            f.write(f"下載失敗，無找到檔案: {datetime.now()}\n")

    # ==== 新增: 如果有下載文件，用 Zoho Mail 發送郵件 ====
    if downloaded_files:
        print("開始發送郵件...")
        try:
            # Zoho Mail SMTP 設定
            smtp_server = 'smtp.zoho.com'
            smtp_port = 587
            sender_email = 'paklun_ckline@zohomail.com'  # 發送人電郵
            sender_password = os.environ.get('ZOHO_PASSWORD', 'default_zoho_password')
            receiver_email = 'paklun@ckline.com.hk'  # 收件人電郵

            # 創建郵件
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = receiver_email
            msg['Subject'] = f"HIT DAILY + {datetime.now().strftime('%Y-%m-%d')}"

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
            print("郵件發送成功!")
        except Exception as e:
            print(f"郵件發送失敗: {str(e)}")
            with open("error_log.txt", "a") as f:
                f.write(f"郵件發送失敗: {str(e)} - {datetime.now()}\n")
    else:
        print("無文件可發送")

    # 設置腳本完成標記
    with open("script_completed.txt", "w") as f:
        f.write(f"腳本完成於: {datetime.now()}\n")
    print("腳本完成標記已創建")

    # 檢查標記是否有效
    if os.path.exists("script_completed.txt"):
        print("確認腳本完成標記存在")
        with open("script_completed.txt", "r") as f:
            print(f"腳本完成時間: {f.read().strip()}")
    else:
        print("腳本完成標記創建失敗")
        with open("error_log.txt", "a") as f:
            f.write(f"腳本完成標記創建失敗 - {datetime.now()}\n")

    # 確保失敗後亦執行登出
    print("嘗試登出...")
    try:
        logout_menu = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button")))
        logout_menu.click()
        time.sleep(2)
        logout_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='menu-list-grow']/div[6]/li")))
        logout_button.click()
        print("登出完成")
    except Exception as logout_error:
        print("登出失敗:", logout_error)
        with open("error_log.txt", "a") as f:
            f.write(f"登出失敗: {str(logout_error)} - {datetime.now()}\n")

except Exception as e:
    print("發生錯誤:", e)
    with open("error_log.txt", "a") as f:
        f.write(f"發生錯誤: {str(e)} - {datetime.now()}\n")
    try:
        print("嘗試緊急登出...")
        try:
            logout_menu = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button")))
            logout_menu.click()
            time.sleep(2)
            logout_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='menu-list-grow']/div[6]/li")))
            logout_button.click()
            print("緊急登出完成")
        except Exception as emergency_logout_error:
            print("緊急登出失敗:", emergency_logout_error)
            with open("error_log.txt", "a") as f:
                f.write(f"緊急登出失敗: {str(emergency_logout_error)} - {datetime.now()}\n")
    except Exception:
        pass

finally:
    driver.quit()
