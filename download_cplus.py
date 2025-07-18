# ... (之前代碼)

# 點擊 Download (Barge)
print("點擊 Download...", flush=True)
download_button_barge = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='content-mount']/app-download-report/div[2]/div/form/div[2]/button")))
download_button_barge.click()
print("Download 按鈕點擊成功", flush=True)
time.sleep(120)  # 延長下載等待時間

# 檢查 Barge Container Detail 下載文件
print("檢查 Barge Container Detail 下載文件...", flush=True)
start_time = time.time()
while time.time() - start_time < 120:
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
while time.time() - start_time < 120:
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
        receiver_email = 'paklun@ckline.com.hk'

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
