import os
import time
import subprocess
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# 設置目錄（現時無需下載）
download_dir = os.path.abspath("housekeep_downloads")
if not os.path.exists(download_dir):
    os.makedirs(download_dir)
    print(f"創建目錄: {download_dir}", flush=True)

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
chrome_options.add_argument('--disable-blink-features=AutomationControlled')
chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
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
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))
    print(f"網站已成功打開，當前 URL: {driver.current_url}", flush=True)

    # 點擊登錄前嘅按鈕
    print("點擊登錄前按鈕...", flush=True)
    wait = WebDriverWait(driver, 10)
    login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
    login_button_pre.click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    print("登錄前按鈕點擊成功", flush=True)

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
    login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/form/div[2]/div/div/form/button/span[1]")
    login_button.click()
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]")))
    print("LOGIN 按鈕點擊成功", flush=True)

    # 前往 housekeepReport 頁面
    print("前往 housekeepReport 頁面...", flush=True)
    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")))  # 延長等待
    print("housekeepReport 頁面加載完成", flush=True)
    time.sleep(5)

    # 關閉可能嘅對話框
    try:
        close_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButton-containedPrimary') and contains(., 'Close')]")))
        close_button.click()
        WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.XPATH, "//div[contains(@class, 'MuiDialog-root')]")))
        print("關閉對話框成功", flush=True)
    except TimeoutException:
        print("無發現對話框，繼續執行", flush=True)

    # 等待報告表格加載
    print("等待報告表格加載...", flush=True)
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/form/div[1]/div[8]/div/div/div[1]/div[2]/div/div/div/table/tbody")))

    # 選擇所有可見報告選項
    print("選擇所有可見報告選項...", flush=True)
    base_xpath = "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/form/div[1]/div[8]/div/div/div[1]/div[2]/div/div/div/table/tbody"
    inputs = wait.until(EC.presence_of_all_elements_located((By.XPATH, f"{base_xpath}//td[6]//input[@class='jss125' and @type='checkbox']")))
    for i, input_element in enumerate(inputs, 1):
        try:
            print(f"嘗試選擇 tr[{i}]...", flush=True)
            if not input_element.is_selected():
                input_element.click()
                print(f"tr[{i}] 選擇成功", flush=True)
            else:
                print(f"tr[{i}] 已選擇，跳過", flush=True)
            time.sleep(1)
        except Exception as e:
            print(f"tr[{i}] 選擇失敗: {str(e)}", flush=True)
            continue

    # 點擊 EMAIL 按鈕
    print("點擊 EMAIL 按鈕...", flush=True)
    try:
        email_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[3]/div/div/form/div[1]/div[8]/div/div/div[1]/div[1]/div[1]/button[1]")), timeout=20)
        email_button.click()
        print("EMAIL 按鈕點擊成功", flush=True)
    except TimeoutException:
        print("主 EMAIL 按鈕未找到，嘗試備用定位...", flush=True)
        email_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButton-contained') and contains(., 'Email')]")), timeout=20)
        email_button.click()
        print("備用 EMAIL 按鈕點擊成功", flush=True)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='to']")))

    # 輸入 TO 字段
    print("輸入收件人 paklun@ckline.com.hk...", flush=True)
    to_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='to']")))
    to_field.clear()
    to_field.send_keys("paklun@ckline.com.hk")
    print("收件人輸入完成", flush=True)
    time.sleep(1)

    # 點擊 CONFIRM 按鈕
    print("點擊 CONFIRM 按鈕...", flush=True)
    confirm_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='EmailDialog']/div[3]/div/form/div[3]/button[1]/span[1]")))
    confirm_button.click()
    print("CONFIRM 按鈕點擊成功", flush=True)
    WebDriverWait(driver, 15).until(EC.invisibility_of_element_located((By.XPATH, "//*[@id='EmailDialog']")))  # 等待對話框消失

    # 點擊工具欄進行登出
    print("點擊工具欄進行登出...", flush=True)
    logout_toolbar = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-toolbar']/button[4]/span[1]")))
    driver.execute_script("arguments[0].scrollIntoView(true);", logout_toolbar)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", logout_toolbar)
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

    print("腳本完成", flush=True)

except Exception as e:
    print(f"發生錯誤: {str(e)}", flush=True)
    raise

finally:
    driver.quit()
