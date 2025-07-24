import os
import time
from datetime import datetime
import pytz
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# 設置 Chrome 選項
def get_chrome_options():
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # 使用舊版 Headless 模式，對齊 download_cplus.py
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--disable-popup-blocking')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--no-first-run')
    chrome_options.add_argument('--window-size=1920,1080')  # 設置窗口大小
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')  # 繞過自動化檢測
    chrome_options.binary_location = '/snap/bin/chromium'
    return chrome_options

# 主任務邏輯
def process_download_housekeep():
    driver = None
    try:
        # 配置 chromedriver
        chromedriver_path = '/home/runner/chromium-bin/chromedriver'
        print(f"Download Housekeep: 使用 chromedriver 路徑: {chromedriver_path}", flush=True)
        if not os.path.exists(chromedriver_path):
            raise FileNotFoundError(f"chromedriver 未找到: {chromedriver_path}")
        service = Service(chromedriver_path)
        driver = webdriver.Chrome(service=service, options=get_chrome_options())
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
        print("
