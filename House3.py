import os
import time
import shutil
import subprocess
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, WebDriverException
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

cplus_download_dir = os.path.abspath("downloads_cplus")
MAX_RETRIES = 3
DOWNLOAD_TIMEOUT = 15

def clear_download_dirs():
    for dir_path in [cplus_download_dir]:
        if os.path.exists(dir_path):
            shutil.rmtree(dir_path)
        os.makedirs(dir_path)
        logging.info(f"創建下載目錄: {dir_path}")

def setup_environment():
    try:
        result = subprocess.run(['which', 'chromium-browser'], capture_output=True, text=True)
        if result.returncode != 0:
            subprocess.run(['sudo', 'apt-get', 'update', '-qq'], check=True)
            subprocess.run(['sudo', 'apt-get', 'install', '-y', 'chromium-browser', 'chromium-chromedriver'], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            logging.info("Chromium 及 ChromeDriver 已安裝")
        else:
            logging.info("Chromium 及 ChromeDriver 已存在，跳過安裝")

        result = subprocess.run(['pip', 'show', 'selenium'], capture_output=True, text=True)
        if "selenium" not in result.stdout:
            subprocess.run(['pip', 'install', 'selenium'], check=True)
            logging.info("Selenium 已安裝")
        else:
            logging.info("Selenium 已存在，跳過安裝")
    except subprocess.CalledProcessError as e:
        logging.error(f"環境準備失敗: {e}")
        raise

def get_chrome_options(download_dir):
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--disable-popup-blocking')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--no-first-run')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": False
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.binary_location = '/usr/bin/chromium-browser'
    return chrome_options

def wait_for_new_file(download_dir, initial_files, timeout=DOWNLOAD_TIMEOUT):
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_files = set(os.listdir(download_dir))
        new_files = current_files - initial_files
        if new_files and any(os.path.getsize(os.path.join(download_dir, f)) > 0 for f in new_files if f.endswith(('.csv', '.xlsx'))):
            return new_files, time.time() - start_time
        time.sleep(0.1)
    return set(), 0

def handle_popup(driver, wait):
    max_attempts = 3
    attempt = 0
    while attempt < max_attempts:
        try:
            error_div = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'System Error') or contains(text(), 'Loading') or contains(@class, 'error') or contains(@class, 'popup')]")))
            logging.error(f"CPLUS: 檢測到系統錯誤: {error_div.text}")
            close_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Close') or contains(text(), 'OK') or contains(text(), 'Download Screenshots')]")))
            driver.execute_script("arguments[0].click();", close_button)
            logging.info("CPLUS: 關閉系統錯誤彈窗")
            time.sleep(1)
            driver.save_screenshot("system_error.png")
            with open("system_error.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            attempt += 1
        except TimeoutException:
            logging.debug("無系統錯誤或加載彈窗檢測到")
            break

def cplus_login(driver, wait):
    logging.info("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/")
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    logging.info(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}")
    time.sleep(1)

    logging.info("CPLUS: 點擊登錄前按鈕...")
    login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
    ActionChains(driver).move_to_element(login_button_pre).click().perform()
    logging.info("CPLUS: 登錄前按鈕點擊成功")
    time.sleep(0.5)

    logging.info("CPLUS: 輸入 COMPANY CODE...")
    company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    company_code_field.send_keys("CKL")
    logging.info("CPLUS: COMPANY CODE 輸入完成")
    time.sleep(0.5)

    logging.info("CPLUS: 輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
    user_id_field.send_keys("KEN")
    logging.info("CPLUS: USER ID 輸入完成")
    time.sleep(0.5)

    logging.info("CPLUS: 輸入 PASSWORD...")
    password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
    password_field.send_keys(os.environ.get('SITE_PASSWORD'))
    logging.info("CPLUS: PASSWORD 輸入完成")
    time.sleep(0.5)

    logging.info("CPLUS: 點擊 LOGIN 按鈕...")
    login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
    ActionChains(driver).move_to_element(login_button).click().perform()
    logging.info("CPLUS: LOGIN 按鈕點擊成功")
    try:
        handle_popup(driver, wait)
        WebDriverWait(driver, 30).until(EC.url_contains("https://cplus.hit.com.hk/app/#/"))
        logging.info("CPLUS: 檢測到主界面 URL，登錄成功")
    except TimeoutException:
        logging.warning("CPLUS: 登錄後主界面加載失敗，但繼續執行")
        driver.save_screenshot("login_timeout.png")
        with open("login_timeout.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)

def attempt_click(button, driver):
    methods = [
        ("Standard click", lambda: button.click()),
        ("ActionChains click", lambda: ActionChains(driver).move_to_element(button).click().perform()),
        ("ActionChains click_and_hold", lambda: ActionChains(driver).move_to_element(button).click_and_hold().pause(0.1).release().perform()),
        ("JavaScript click", lambda: driver.execute_script("arguments[0].click();", button)),
        ("Send Keys ENTER", lambda: ActionChains(driver).move_to_element(button).send_keys(Keys.ENTER).perform()),
        ("DispatchEvent click", lambda: driver.execute_script("arguments[0].dispatchEvent(new Event('click'));", button))
    ]
    results = {method_name: False for method_name, _ in methods}
    for method_name, method_func in methods:
        try:
            method_func()
            results[method_name] = True
            logging.debug(f"點擊測試方法 {method_name} 成功")
        except Exception as e:
            logging.debug(f"點擊測試方法 {method_name} 失敗: {str(e)}")
    successful = [name for name, success in results.items() if success]
    return bool(successful), successful[0] if successful else None

def process_cplus_house(driver, wait, initial_files):
    logging.info("CPLUS: 前往 Housekeeping Reports 頁面...")
    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")), 60)
        logging.info("CPLUS: Housekeeping Reports 頁面加載完成")
    except TimeoutException:
        logging.error("CPLUS: House 頁面加載失敗，刷新頁面...")
        driver.refresh()
        time.sleep(5)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")), 60)
        logging.info("CPLUS: Housekeeping Reports 頁面加載完成 (刷新後)")

    logging.info("CPLUS: 等待表格加載...")
    start_time = time.time()
    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//table")))
        rows = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, "//table//tbody//tr")))
        if len(rows) == 0:
            logging.warning("CPLUS: 無記錄，嘗試刷新...")
            driver.refresh()
            time.sleep(5)
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//table")))
            rows = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, "//table//tbody//tr")))
        logging.info(f"CPLUS: 找到 {len(rows)} 個報告行，耗時 {time.time() - start_time:.1f} 秒")
    except TimeoutException:
        logging.warning("CPLUS: 表格加載失敗，嘗試備用定位...")
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
        rows = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table tbody tr")))
        if len(rows) == 0:
            logging.error("CPLUS: 表格加載完全失敗，跳過 Housekeeping Reports")
            driver.save_screenshot("house_load_failure.png")
            with open("house_load_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            return set(), 0, 0
        logging.info(f"CPLUS: 找到 {len(rows)} 個報告行 (備用定位)，耗時 {time.time() - start_time:.1f} 秒")

    if time.time() - start_time > 60:
        logging.warning("CPLUS: Housekeeping Reports 加載時間過長，跳過")
        driver.save_screenshot("house_load_timeout.png")
        return set(), 0, 0

    time.sleep(0.5)
    logging.info("CPLUS: 定位並點擊所有 EXCEL 下載按鈕...")
    local_initial = initial_files.copy()
    new_files = set()
    all_downloaded_files = set()
    excel_buttons = driver.find_elements(By.XPATH, "//table//tbody//tr//td//button[contains(text(), 'EXCEL')]")  # 使用 contains(text()) 確保文本匹配
    button_count = len(excel_buttons)
    if button_count == 0:
        logging.error("CPLUS: 未找到任何 EXCEL 下載按鈕，記錄頁面狀態...")
        driver.save_screenshot("house_button_failure.png")
        with open("house_button_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: Housekeeping Reports 未找到 EXCEL 下載按鈕")
    logging.info(f"CPLUS: 找到 {button_count} 個 EXCEL 下載按鈕 (預期 6 個)")

    # 記錄實際按鈕文本以 debug
    for idx, btn in enumerate(excel_buttons, 1):
        btn_text = btn.text or btn.get_attribute("innerText") or "無文本"
        logging.debug(f"按鈕 {idx} 文本: {btn_text}")

    # 驗證按鈕數量
    if button_count != 6:
        logging.warning(f"CPLUS: 檢測到 {button_count} 個 EXCEL 按鈕，預期 6 個，繼續執行但可能有誤")

    report_file_mapping = []
    failed_buttons = []
    click_times = []
    successful_methods = {
        "Standard click": 0,
        "ActionChains click": 0,
        "ActionChains click_and_hold": 0,
        "JavaScript click": 0,
        "Send Keys ENTER": 0,
        "DispatchEvent click": 0
    }

    for idx in range(button_count):
        success = False
        for retry_count in range(MAX_RETRIES):
            try:
                button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, f"(//table//tbody//tr//td//button[contains(text(), 'EXCEL')])[{idx+1}]")))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", button)
                click_time = time.time()
                click_times.append(click_time)
                time.sleep(0.1)

                try:
                    report_name = driver.find_element(By.XPATH, f"(//table//tbody//tr//td[3])[{idx+1}]").text
                    logging.info(f"CPLUS: 準備點擊第 {idx+1} 個 EXCEL 按鈕，報告名稱: {report_name}")
                except Exception as e:
                    report_name = f"未知報告 {idx+1}"
                    logging.warning(f"CPLUS: 獲取報告名稱失敗: {str(e)}")

                clicked, method_names = attempt_click(button, driver)
                if not clicked:
                    raise Exception("所有點擊方法失敗")
                for method in method_names:
                    successful_methods[method] += 1

                handle_popup(driver, wait)
                time.sleep(1)

                temp_new, download_time = wait_for_new_file(cplus_download_dir, local_initial)
                if temp_new:
                    matched_file = next(f for f in temp_new if f.endswith(('.csv', '.xlsx')))
                    all_downloaded_files.add(matched_file)
                    report_file_mapping.append((report_name, matched_file, download_time))
                    local_initial.add(matched_file)
                    new_files.add(matched_file)
                    success = True
                    logging.info(f"CPLUS: 第 {idx+1} 個下載成功，文件: {matched_file}, 耗時 {download_time:.1f} 秒，使用方法: {', '.join(method_names)}")
                    break
                else:
                    logging.warning(f"CPLUS: 第 {idx+1} 個未觸發新 EXCEL 文件 (重試 {retry_count+1})")
                    driver.save_screenshot(f"house_button_{idx+1}_failure_{retry_count}.png")
                    with open(f"house_button_{idx+1}_failure_{retry_count}.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
            except (TimeoutException, NoSuchElementException, ElementClickInterceptedException, WebDriverException) as e:
                logging.error(f"CPLUS: 第 {idx+1} 個失敗: {str(e)}")
                driver.save_screenshot(f"house_button_{idx+1}_failure_{retry_count}.png")
                if retry_count < MAX_RETRIES - 1:
                    logging.info(f"CPLUS: 刷新頁面重試第 {idx+1} 個...")
                    driver.refresh()
                    time.sleep(5)
        if not success:
            failed_buttons.append(idx)
            report_file_mapping.append((report_name, "N/A", 0))

    if new_files:
        logging.info(f"CPLUS: Housekeeping Reports 下載完成，共 {len(new_files)} 個文件，找到 {button_count} 個 EXCEL 按鈕")
        for report, files, _ in report_file_mapping:
            logging.info(f"報告: {report}, 文件: {files}")
        if failed_buttons:
            logging.warning(f"CPLUS: {len(failed_buttons)} 個失敗: {failed_buttons}")
    else:
        logging.warning("CPLUS: Housekeeping Reports 無任何 EXCEL 下載")

    logging.info("點擊方法穩定性統計:")
    for method, count in successful_methods.items():
        logging.info(f"{method}: {count} 次成功")

    return new_files, len(new_files), button_count

def main():
    load_dotenv()
    clear_download_dirs()
    setup_environment()
    driver = None
    initial_files = set(os.listdir(cplus_download_dir))
    try:
        driver = webdriver.Chrome(options=get_chrome_options(cplus_download_dir))
        logging.info("CPLUS WebDriver 初始化成功")
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        wait = WebDriverWait(driver, 30)
        cplus_login(driver, wait)
        new_files, house_file_count, house_button_count = process_cplus_house(driver, wait, initial_files)
        logging.info(f"總下載文件: {len(new_files)} 個")
        for file in new_files:
            logging.info(f"找到檔案: {file}")
    finally:
        if driver:
            try:
                driver.quit()
                logging.info("CPLUS WebDriver 關閉")
            except Exception as e:
                logging.error(f"CPLUS WebDriver 關閉失敗: {str(e)}")

if __name__ == "__main__":
    main()
