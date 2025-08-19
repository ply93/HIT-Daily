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
MAX_RETRIES = 1  # 保持 1 次
DOWNLOAD_TIMEOUT = 3  # 保持 3 秒

def clear_download_dirs():
    for dir_path in [cplus_download_dir]:
        if os.path.exists(dir_path):
            for file in os.listdir(dir_path):
                file_path = os.path.join(dir_path, file)
                if os.path.isfile(file_path):
                    os.remove(file_path)
            os.makedirs(dir_path, exist_ok=True)
        else:
            os.makedirs(dir_path)
        logging.info(f"創建並清空下載目錄: {dir_path}")

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
    chrome_options.set_capability('timeouts', {'implicit': 30000, 'pageLoad': 30000, 'script': 30000})  # 保持 30 秒
    return chrome_options

def wait_for_new_file(download_dir, initial_files, expected_filename=None, timeout=DOWNLOAD_TIMEOUT):
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_files = set(os.listdir(download_dir))
        new_files = current_files - initial_files
        if new_files:
            for file in new_files:
                file_path = os.path.join(download_dir, file)
                if os.path.getsize(file_path) > 0 and file.endswith(('.csv', '.xlsx')):
                    logging.debug(f"檢測到新文件: {file}, 預期: {expected_filename}")
                    return {file}, time.time() - start_time
        time.sleep(0.1)
    logging.warning(f"下載超時，當前文件: {list(current_files)}")
    return set(), 0

def handle_popup(driver, wait):
    max_attempts = 3
    attempt = 0
    while attempt < max_attempts:
        try:
            error_div = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'System Error') or contains(text(), 'Loading') or contains(@class, 'error') or contains(@class, 'popup')]")))  # 保持 5 秒
            logging.error(f"CPLUS: 檢測到系統錯誤: {error_div.text}")
            close_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Close') or contains(text(), 'OK') or contains(text(), 'Download Screenshots')]")))
            driver.execute_script("arguments[0].click();", close_button)
            logging.info("CPLUS: 關閉系統錯誤彈窗")
            time.sleep(0.2)  # 保持 0.2 秒
            driver.save_screenshot("system_error.png")
            with open("system_error.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            attempt += 1
        except TimeoutException:
            logging.debug("無系統錯誤或加載彈窗檢測到，繼續執行")
            break
        except Exception as e:
            logging.warning(f"彈窗處理失敗: {str(e)}，嘗試下一個...")
            attempt += 1
            time.sleep(0.2)

def cplus_login(driver, wait):
    logging.info("CPLUS: 嘗試打開網站 https://cplus.hit.com.hk/frontpage/#/")
    driver.get("https://cplus.hit.com.hk/frontpage/#/")
    logging.info(f"CPLUS: 網站已成功打開，當前 URL: {driver.current_url}")
    time.sleep(0.2)  # 保持 0.2 秒

    logging.info("CPLUS: 點擊登錄前按鈕...")
    login_button_pre = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/button/span[1]")))
    ActionChains(driver).move_to_element(login_button_pre).click().perform()
    logging.info("CPLUS: 登錄前按鈕點擊成功")
    time.sleep(0.2)

    logging.info("CPLUS: 輸入 COMPANY CODE...")
    company_code_field = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='companyCode']")))
    company_code_field.send_keys("CKL")
    logging.info("CPLUS: COMPANY CODE 輸入完成")
    time.sleep(0.2)

    logging.info("CPLUS: 輸入 USER ID...")
    user_id_field = driver.find_element(By.XPATH, "//*[@id='userId']")
    user_id_field.send_keys("KEN")
    logging.info("CPLUS: USER ID 輸入完成")
    time.sleep(0.2)

    logging.info("CPLUS: 輸入 PASSWORD...")
    password_field = driver.find_element(By.XPATH, "//*[@id='passwd']")
    password_field.send_keys(os.environ.get('SITE_PASSWORD'))
    logging.info("CPLUS: PASSWORD 輸入完成")
    time.sleep(0.2)

    logging.info("CPLUS: 點擊 LOGIN 按鈕...")
    login_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/header/div/div[4]/div[2]/div/div/form/button/span[1]")
    ActionChains(driver).move_to_element(login_button).click().perform()
    logging.info("CPLUS: LOGIN 按鈕點擊成功")
    for _ in range(3):  # 增加重試邏輯
        try:
            handle_popup(driver, wait)
            WebDriverWait(driver, 5).until(EC.url_contains("https://cplus.hit.com.hk/app/#/"))  # 保持 5 秒
            logging.info("CPLUS: 檢測到主界面 URL，登錄成功")
            break
        except TimeoutException:
            logging.warning("CPLUS: 登錄後主界面加載失敗，重試...")
            time.sleep(1)
        except Exception as e:
            logging.error(f"CPLUS: 登錄失敗: {str(e)}，重試...")
            time.sleep(1)
    else:
        logging.error("CPLUS: 登錄重試失敗，記錄頁面狀態...")
        driver.save_screenshot("login_failure.png")
        with open("login_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: 登錄失敗")

def attempt_click(button, driver, method_name):
    methods = {
        "Standard click": lambda: button.click(),
        "ActionChains click": lambda: ActionChains(driver).move_to_element(button).click().perform(),
        "ActionChains click_and_hold": lambda: ActionChains(driver).move_to_element(button).click_and_hold().pause(0.1).release().perform(),
        "JavaScript click": lambda: driver.execute_script("arguments[0].click();", button),
        "Send Keys ENTER": lambda: ActionChains(driver).move_to_element(button).send_keys(Keys.ENTER).perform(),
        "DispatchEvent click": lambda: driver.execute_script("arguments[0].dispatchEvent(new Event('click'));", button)
    }
    try:
        if not button.is_enabled() or not button.is_displayed():
            logging.warning(f"按鈕不可點擊: {method_name}")
            return False
        methods[method_name]()
        logging.debug(f"點擊測試方法 {method_name} 成功")
        return True
    except WebDriverException as e:
        logging.error(f"WebDriverException 發生: {str(e)}，方法: {method_name}")
        return False
    except Exception as e:
        logging.debug(f"點擊測試方法 {method_name} 失敗: {str(e)}")
        return False

def process_cplus_house(driver, wait, initial_files):
    logging.info("CPLUS: 前往 Housekeeping Reports 頁面...")
    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")), 5)  # 保持 5 秒
        logging.info("CPLUS: Housekeeping Reports 頁面加載完成")
    except TimeoutException:
        logging.error("CPLUS: House 頁面加載失敗，刷新頁面...")
        driver.refresh()
        time.sleep(2)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']")), 5)
        logging.info("CPLUS: Housekeeping Reports 頁面加載完成 (刷新後)")

    logging.info("CPLUS: 等待表格加載...")
    start_time = time.time()
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]")))  # 保持 5 秒
        rows = WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr[td[3]]")))  # 匹配有報告名稱的行
        logging.info(f"CPLUS: 找到 {len(rows)} 個報告行，耗時 {time.time() - start_time:.1f} 秒")
    except TimeoutException:
        logging.warning("CPLUS: 表格加載失敗，嘗試備用定位...")
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
        rows = WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table tbody tr")))
        logging.info(f"CPLUS: 找到 {len(rows)} 個報告行 (備用定位)，耗時 {time.time() - start_time:.1f} 秒")

    if time.time() - start_time > 10:  # 保持 10 秒
        logging.warning("CPLUS: Housekeeping Reports 加載時間過長，跳過")
        driver.save_screenshot("house_load_timeout.png")
        return set(), 0, 0

    time.sleep(0.2)  # 保持 0.2 秒
    logging.info("CPLUS: 定位並點擊所有 Excel 下載按鈕...")
    local_initial = initial_files.copy()
    new_files = set()
    all_downloaded_files = set()
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)]")))  # 參考第一種定位
        excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)]")
        button_count = len(excel_buttons)
        logging.info(f"CPLUS: 找到 {button_count} 個 Excel 下載按鈕")
        if button_count == 0:
            logging.debug("CPLUS: 未找到 Excel 按鈕，嘗試原始定位...")
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]//svg[@viewBox='0 0 24 24']//path[@fill='#036e11']")))  # 參考第二種定位
            excel_buttons = driver.find_elements(By.XPATH, "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]//svg[@viewBox='0 0 24 24']//path[@fill='#036e11']")
            button_count = len(excel_buttons)
            logging.info(f"CPLUS: 原始定位找到 {button_count} 個 Excel 下載按鈕")
    except TimeoutException:
        logging.error("CPLUS: 按鈕加載失敗，記錄頁面狀態...")
        driver.save_screenshot("button_load_failure.png")
        with open("button_load_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: Housekeeping Reports 按鈕加載失敗")

    if button_count == 0:
        logging.error("CPLUS: 未找到任何 Excel 下載按鈕，記錄頁面狀態...")
        driver.save_screenshot("house_button_failure.png")
        with open("house_button_failure.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("CPLUS: Housekeeping Reports 未找到 Excel 下載按鈕")
    logging.info(f"CPLUS: 最終找到 {button_count} 個 Excel 下載按鈕")

    # 記錄實際按鈕屬性以 debug
    for idx, btn in enumerate(excel_buttons, 1):
        btn_text = btn.text or btn.get_attribute("innerText") or btn.get_attribute("title") or btn.get_attribute("aria-label") or "無文本"
        btn_class = btn.get_attribute("class") or "無類別"
        logging.debug(f"按鈕 {idx} 文本/title/aria-label: {btn_text}, 類別: {btn_class}")

    click_methods = [
        "Standard click",
        "ActionChains click",
        "ActionChains click_and_hold",
        "JavaScript click",
        "Send Keys ENTER",
        "DispatchEvent click"
    ]
    successful_methods = {method: 0 for method in click_methods}
    report_file_mapping = []
    failed_buttons = []

    for method in click_methods:
        logging.info(f"CPLUS: 開始測試點擊方法: {method}")
        local_initial = initial_files.copy()  # 重置初始文件集
        for idx in range(button_count):
            success = False
            report_name = driver.find_element(By.XPATH, f"(//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[3])[{idx+1}]").text
            expected_filename = f"{report_name.replace(' ', '_').replace('/', '_')}_{time.strftime('%d%m%y')}_CKL"
            for retry_count in range(MAX_RETRIES):
                try:
                    button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"(//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]/div/button[not(@disabled)])[{idx+1}]" if idx < len(excel_buttons) else "//table[contains(@class, 'MuiTable-root')]//tbody//tr//td[4]//button[not(@disabled)]//svg[@viewBox='0 0 24 24']//path[@fill='#036e11']")))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", button)
                    click_time = time.time()
                    time.sleep(0.2)  # 保持 0.2 秒

                    logging.info(f"CPLUS: 準備點擊第 {idx+1} 個 EXCEL 按鈕，報告名稱: {report_name}，使用方法: {method}")
                    clicked = attempt_click(button, driver, method)
                    if clicked:
                        successful_methods[method] += 1
                        logging.info(f"成功點擊方法: {method}")
                    else:
                        raise Exception(f"點擊方法 {method} 失敗")

                    handle_popup(driver, wait)
                    time.sleep(0.2)

                    temp_new, download_time = wait_for_new_file(cplus_download_dir, local_initial)
                    if temp_new:
                        matched_file = temp_new.pop()
                        all_downloaded_files.add(matched_file)
                        report_file_mapping.append((report_name, matched_file, download_time))
                        local_initial.add(matched_file)
                        new_files.add(matched_file)
                        success = True
                        logging.info(f"CPLUS: 第 {idx+1} 個下載成功，文件: {matched_file}, 預期: {expected_filename}, 耗時 {download_time:.1f} 秒，使用方法: {method}")
                        break
                    else:
                        logging.warning(f"CPLUS: 第 {idx+1} 個未觸發新 EXCEL 文件 (重試 {retry_count+1})，使用方法: {method}")
                        driver.save_screenshot(f"house_button_{idx+1}_failure_{method}_{retry_count}.png")
                        with open(f"house_button_{idx+1}_failure_{method}_{retry_count}.html", "w", encoding="utf-8") as f:
                            f.write(driver.page_source)
                except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
                    logging.error(f"CPLUS: 第 {idx+1} 個失敗: {str(e)}，使用方法: {method}")
                    driver.save_screenshot(f"house_button_{idx+1}_failure_{method}_{retry_count}.png")
                    if retry_count < MAX_RETRIES - 1:
                        logging.info(f"CPLUS: 刷新頁面重試第 {idx+1} 個...")
                        driver.refresh()
                        time.sleep(2)
                except WebDriverException as e:
                    logging.error(f"WebDriver 異常: {str(e)}，關閉並重試...")
                    driver.quit()
                    driver = webdriver.Chrome(options=get_chrome_options(cplus_download_dir))
                    wait = WebDriverWait(driver, 5)
                    cplus_login(driver, wait)
                    driver.get("https://cplus.hit.com.hk/app/#/report/housekeepReport")
                    time.sleep(2)
            if not success:
                failed_buttons.append((idx, method))
        logging.info(f"CPLUS: 點擊方法 {method} 測試完成，成功下載: {sum(1 for r, f, _ in report_file_mapping if f != 'N/A' and method in f)} 個")

    if new_files:
        logging.info(f"CPLUS: Housekeeping Reports 下載完成，共 {len(new_files)} 個文件，找到 {button_count} 個 EXCEL 按鈕")
        for report, files, _ in report_file_mapping:
            logging.info(f"報告: {report}, 文件: {files}")
        if failed_buttons:
            logging.warning(f"CPLUS: 失敗按鈕: {failed_buttons}")
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
        wait = WebDriverWait(driver, 5)
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
