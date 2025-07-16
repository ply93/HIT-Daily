from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(options=options)
driver.get("https://colab.research.google.com/drive/1BbJxRsYnrlHi_RN9zspiErEh3NAf9n83?usp=drive_link")

try:
    cookies = pickle.load(open("cookies.pkl", "rb"))
    for cookie in cookies:
        driver.add_cookie(cookie)
    driver.refresh()
    run_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "run-button"))
    )
    run_button.click()
    print("Clicked Run button")
    time.sleep(10)
except Exception as e:
    print(f"Error: {e}")
finally:
    driver.quit()
