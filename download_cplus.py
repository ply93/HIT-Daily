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

# 設置下載目錄
download_dir = "downloads"
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
prefs = {"download.default_directory": download_dir, "download.prompt_for_download": False}
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
    # 前往登入頁面
    print("嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/", flush=True)
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    print(f"網站已成功打開，當前 URL: {driver.current_url}", flush=True)
    time.sleep(2)

    # 點擊登錄前嘅按鈕
    print("點擊登錄前按鈕...", flush=True)
    wait = WebDriverWait(driver, 20)
    login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
    login_button_pre.click()
    print("登錄前按鈕點擊成功", flush=True)
    time.sleep(2)

    # 輸入 COMPANY CODE
    print("輸入 COMPANY CODE...", flush=True)
    company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    company_code_field.send_keys("CKL")
    print("COMPANY CODE 輸入完成", flush=True)
    time.sleep(1)

    # 輸入 USER ID
    print("輸入 USER ID...", flush=True)
    user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
    user_id_field.send_keys("KEN")
    print("USER ID 輸入完成", flush=True)
    time.sleep(1)

    # 輸入 PASSWORD
    print("輸入 PASSWORD...", flush=True)
    password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
    password_field.send_keys(os.environ.get('SITE_PASSWORD', 'Ken2807890'))
    print("PASSWORD 輸入完成", flush=True)
    time.sleep(1)

    # 點擊 LOGIN 按鈕
    print("點擊 LOGIN 按鈕...", flush=True)
    login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
    login_button.click()
    print("LOGIN 按鈕點擊成功", flush=True)
    time.sleep(15)

    # 前往 Container Movement Log 頁面
    print("直接前往 Container Movement Log...", flush=True)
    driver.get("https://cplus.hit.com.hk/app/#/enquiry/ContainerMovementLog")
    time.sleep(10)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    print("Container Movement Log 頁面加載完成", flush=True)
    time.sleep(5)

    # 點擊 Search
    print("點擊 Search...", flush=True)
    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[1]/div/form/div[2]/div/div[4]/button/span[1]")))
    search_button.click()
    print("Search 按鈕點擊成功", flush=True)
    time.sleep(15)

    # 點擊 Download
    print("點擊 Download...", flush=True)
    download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div[3]/div/div[2]/div/div[2]/div/div[1]/div[1]/button")))
    download_button.click()
    print("Download 按鈕點擊成功", flush=True)
    time.sleep(60)

    # 檢查下載文件
    print("檢查下載文件...", flush=True)
    downloaded_files = [f for f in os.listdir(download_dir) if f.endswith(('.csv', '.xlsx'))]
    if downloaded_files:
        print(f"下載完成，檔案位於: {download_dir}", flush=True)
        for file in downloaded_files:
            print(f"找到檔案: {file}", flush=True)
    else:
        print("下載失敗，無找到檔案", flush=True)

    print("腳本完成", flush=True)

except Exception as e:
    print(f"發生錯誤: {str(e)}", flush=True)
    raise

finally:
    driver.quit()
