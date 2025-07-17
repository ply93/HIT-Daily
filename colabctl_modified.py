import sys
sys.path.insert(0, '/usr/lib/chromium-browser/chromedriver')
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import pickle
import time
import validators
import os

def sleep(seconds):
    for i in range(seconds):
        try:
            time.sleep(1)
        except KeyboardInterrupt:
            continue

def exists_by_text2(driver, text):
    try:
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '{}')]".format(str(text)))))
    except Exception:
        return False
    return True

def exists_by_xpath(driver, thex, howlong):
    try:
        WebDriverWait(driver, howlong).until(EC.visibility_of_element_located((By.XPATH, thex)))
    except:
        return False

def exists_by_text(driver, text):
    driver.implicitly_wait(2)
    try:
        driver.find_element(By.XPATH, "//*[contains(text(), '{}')]".format(str(text)))
    except NoSuchElementException:
        driver.implicitly_wait(5)
        return False
    driver.implicitly_wait(5)
    return True

def user_logged_in(driver):
    try:
        driver.find_element(By.XPATH, '//*[@id="file-type"]')
    except NoSuchElementException:
        driver.implicitly_wait(5)
        return False
    driver.implicitly_wait(5)
    return True

def wait_for_xpath(driver, x):
    while True:
        try:
            driver.find_element(By.XPATH, x)
            return True
        except:
            time.sleep(0.1)
            pass

def scroll_to_bottom(driver):
    SCROLL_PAUSE_TIME = 0.5
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(SCROLL_PAUSE_TIME)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

def file_to_list(filename):
    colabs = []
    for line in open(filename):
        if validators.url(line):
            colabs.append(line)
    return colabs

def switch_to_tab(driver, tab_index):
    print("Switching to tab " + str(tab_index))
    try:
        driver.switch_to.window(driver.window_handles[tab_index])
    except:
        print("Error switching tabs.")
        return False

def new_tab(driver, url, tab_index):
    print("Opening new tab to " + str(url))
    try:
        driver.execute_script("window.open('" + str(url) + "')")
    except:
        print("Error opening new tab.")
        return False
    switch_to_tab(driver, tab_index)
    return True

fork = sys.argv[1]
if len(sys.argv) > 2 and sys.argv[1] == 'run':
    timeout = int(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[2].isdigit() else 60
    colab_urls = [sys.argv[3]] if len(sys.argv) > 3 else file_to_list('notebooks.csv')

if len(colab_urls) > 0 and validators.url(colab_urls[0]):
    colab_1 = colab_urls[0]
else:
    raise Exception('No notebooks')

chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
wd = webdriver.Chrome('chromedriver', options=chrome_options)

try:
    print("Attempting to log in to Google...")
    wd.get("https://accounts.google.com")
    time.sleep(5)
    email_field = WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, "identifierId")))
    email_field.send_keys(os.environ.get('GOOGLE_EMAIL', 'your_email@gmail.com'))
    email_field.send_keys(Keys.RETURN)
    time.sleep(10)  # 延長等待密碼頁加載
    # 檢查是否進入密碼頁
    try:
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, "passwordNext")))
        password_field = WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.NAME, "Passwd")))
        password_field.send_keys(os.environ.get('GOOGLE_PASSWORD', 'your_password'))
        password_field.send_keys(Keys.RETURN)
        time.sleep(15)  # 延長等待登入完成
        print("Google login completed.")
    except TimeoutException as e:
        print(f"Failed to locate password field: {e}. Possible CAPTCHA or verification required.")
        raise
except Exception as e:
    print(f"Login failed: {e}")
    raise

    wd.get(colab_1)
    try:
        for cookie in pickle.load(open("gCookies.pkl", "rb")):
            wd.add_cookie(cookie)
    except Exception:
        pass
    wd.get(colab_1)

    if exists_by_text(wd, "Sign in"):
        print("No auth cookie detected. Please login to Google.")
        wd.close()
        wd.quit()
        chrome_options_gui = Options()
        chrome_options_gui.add_argument('--no-sandbox')
        chrome_options_gui.add_argument('--disable-infobars')
        user_data_dir = os.path.join(os.getcwd(), "temp_user_data_" + str(time.time()))  # 唯一用戶數據目錄
        chrome_options_gui.add_argument(f'--user-data-dir={user_data_dir}')
        service = Service('/usr/bin/chromedriver')  # 指定 ChromeDriver 路徑
        wd = webdriver.Chrome(service=service, options=chrome_options_gui)
        wd.get("https://accounts.google.com/signin")
        wait_for_xpath(wd, '//*[@id="yDmH0d"]/c-wiz/div/div[2]/c-wiz/c-wiz/div/div[4]/div/div/header/div[2]')
        print("Login detected. Saving cookie & restarting connection.")
        pickle.dump(wd.get_cookies(), open("gCookies.pkl", "wb"))
        wd.close()
        wd.quit()
        wd = webdriver.Chrome('chromedriver', options=chrome_options)

    while True:
        for colab_url in colab_urls:
            complete = False
            wd.get(colab_url)
            print("Logged in.")  # for debugging
            running = False
            wait_for_xpath(wd, '//*[@id="file-menu-button"]/div/div/div[1]')
            print('Notebook loaded.')
            sleep(10)

            while not exists_by_text(wd, "Sign in"):
                if exists_by_text(wd, "Runtime disconnected"):
                    try:
                        wd.find_element_by_xpath('//*[@id="ok"]').click()
                    except NoSuchElementException:
                        pass
                if exists_by_text2(wd, "Notebook loading error"):
                    wd.get(colab_url)
                try:
                    wd.find_element_by_xpath('//*[@id="file-menu-button"]/div/div/div[1]')
                    if not running:
                        wd.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.SHIFT + "q")
                        wd.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.SHIFT + "k")
                        exists_by_xpath(wd, '//*[@id="ok"]', 10)
                        wd.find_element_by_xpath('//*[@id="ok"]').click()
                        sleep(10)
                        wd.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.F9)
                        running = True
                except NoSuchElementException:
                    pass
                if running:
                    try:
                        wd.find_element_by_css_selector('.notebook-content-background').click()
                        scroll_to_bottom(wd)
                        print("performed scroll")
                    except:
                        pass
                    for frame in wd.find_elements_by_tag_name('iframe'):
                        wd.switch_to.frame(frame)
                        for output in wd.find_elements_by_tag_name('pre'):
                            if fork in output.text:
                                running = False
                                complete = True
                                print("Completion string found. Waiting for next cycle.")
                                break
                        wd.switch_to.default_content()
                        if complete:
                            break
                if complete:
                    break
        sleep(timeout)
except Exception as e:
    print(f"Error occurred: {e}")
