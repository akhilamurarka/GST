import pandas as pd
import os, time, requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException, NoSuchElementException
from pydub import AudioSegment
from pydub.utils import which
from word_extractor import transcribe_wav_to_digits
from config import month_name_to_number,quarter_text_map,get_quarter_from_month,month_small_to_big,TIMEOUT
import smtplib
from email.message import EmailMessage
import zipfile

# Set ffmpeg and ffprobe paths manually
AudioSegment.converter = which("ffmpeg")
AudioSegment.ffprobe = which("ffprobe")

def wait_for_new_file(download_dir, before_files, timeout=30):
    """
    Waits for a new file to appear in download_dir compared to before_files.
    Returns the name of the new file.
    """
    for _ in range(timeout):
        current_files = set(os.listdir(download_dir))
        new_files = current_files - before_files
        for file in new_files:
            if not file.endswith(".crdownload"):  # skip temp files
                return file
        time.sleep(1)
    return None

def Chrome_setup(download_path):
    options = Options()
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True,
    }
    options.add_argument("--start-maximized")
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=options)

def Captcha_solving(driver, index, wait):
    for attempt in range(3):
        try:
            audio_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@ng-click='play()']")))
            audio_button.click()
            time.sleep(2)

            audio_element = driver.find_element(By.ID, "audioCap")
            time.sleep(1)
            audio_src = audio_element.get_attribute("src")
            time.sleep(2)
            
            driver.execute_script(f"window.open('{audio_src}', '_blank');")
            time.sleep(1)
            driver.switch_to.window(driver.window_handles[1])
            time.sleep(2)

            session = requests.Session()
            for cookie in driver.get_cookies():
                session.cookies.set(cookie["name"], cookie["value"])

            audio_mp3 = f"captcha_{index+1}.mp3"
            audio_wav = f"captcha_{index+1}.wav"
            response = session.get(audio_src)
            time.sleep(2)
            
            if "audio" not in response.headers.get("Content-Type", ""):
                print("❌ Failed to download audio. Retrying...")
                driver.close()
                time.sleep(1)
                driver.switch_to.window(driver.window_handles[0])
                continue

            with open(audio_mp3, "wb") as f:
                f.write(response.content)
            print("[✔] CAPTCHA audio downloaded")
            
            driver.close()
            time.sleep(1)
            driver.switch_to.window(driver.window_handles[0])

            AudioSegment.from_mp3(audio_mp3).export(audio_wav, format="wav")
            digits = transcribe_wav_to_digits(audio_wav)
            print(f"[🔎 Predicted Digits]: {digits}")
            if len(digits) == 6:
                driver.find_element(By.ID, "captcha").clear()
                time.sleep(1)
                driver.find_element(By.ID, "captcha").send_keys(digits)
                time.sleep(1)
                driver.find_element(By.CSS_SELECTOR, "button.btn.btn-primary").click()
                time.sleep(1)
                return True
            else:
                print("❌ Could not extract 6 digits from audio. Retrying...")
                wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@ng-click='refreshCaptcha()']"))).click()
        except Exception:
            pass
        finally:
            for f in [audio_mp3, audio_wav]:
                if os.path.exists(f): os.remove(f)
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
    return False

def fill_timeline_details(driver, year, quarter_text, month):
    try:
        Select(driver.find_element(By.NAME, "fin")).select_by_visible_text(year)
        time.sleep(2)
        Select(driver.find_element(By.NAME, "quarter")).select_by_visible_text(quarter_text)
        time.sleep(2)
        Select(driver.find_element(By.NAME, "mon")).select_by_visible_text(month)
        time.sleep(2)
        driver.find_element(By.CSS_SELECTOR, "button.btn.btn-primary.srchbtn").click()
        time.sleep(2)
    except Exception as e:
        print(f"[ERROR] Failed to fill timeline details: {e}")
        raise

def send_email_with_attachment(to_email, file_path, sender_email, sender_password, email_subject, email_body):
    # Check if file exists
    if not os.path.exists(file_path):
        print(f"[✘] File not found: {file_path}")
        return False

    try:
        # Create the email message
        msg = EmailMessage()
        msg['Subject'] = email_subject
        msg['From'] = sender_email
        msg['To'] = to_email
        msg.set_content(email_body)

        # Read the file content and add as attachment
        with open(file_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(file_path)

        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

        # Connect to Gmail SMTP server using SSL
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)  # Login with your sender email and app password
            smtp.send_message(msg)

        print(f"[✓] Email sent to {to_email}")
        return True

    except Exception as e:
        print(f"[✘] Failed to send email to {to_email}: {e}")
        return False


def run_automation(excel_path, download_path, progress_callback,sender_email, sender_password, email_subject, email_body):
    df = pd.read_excel(excel_path)
    total = len(df)
    completed = 0

    for index, row in df.iterrows():
        if row['status'] == 'done':
            completed += 1
            continue

        driver = Chrome_setup(download_path)
        wait = WebDriverWait(driver, 10)
        driver.get("https://services.gst.gov.in/services/login")

        wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(row["username"])
        time.sleep(2)
        driver.find_element(By.ID, "user_pass").send_keys(row["password"])
        time.sleep(2)

        if not Captcha_solving(driver, index, wait):
            print(f"[✘] CAPTCHA solving failed for user {row['username']}. Skipping...")
            driver.quit()
            continue
        
        time.sleep(3)
        
        # --- Check for Login Success or Failure ---
        try:
            # Wait up to 10 seconds for a reliable element on the dashboard page.
            # We'll use the 'Return Dashboard' button itself.
            wait.until(
                EC.presence_of_element_located((By.XPATH, "//button[span[@title='Return Dashboard']]"))
            )
            print(f"[✓] Login successful for {row['username']}.")
        
        except TimeoutException:
            # If the dashboard button doesn't appear, the login failed.
            error_message = f"[✘] Login failed for {row['username']}. (Invalid credentials or bad CAPTCHA). Skipping..."

            print(error_message)
            df.at[index, 'status'] = 'login failed' # Mark as failed
            driver.quit()
            
            continue # Move to the next user in the loop
        # --- End of Login Check ---
        
        # Exit any pop-up modal safely, even if it appears twice
        while True:
            try:
                # Try to find and click the cancel popup button
                wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//a[@ng-click='cancelcallback()']"))
                ).click()
                time.sleep(3)

            except TimeoutException:
                # No popup appeared OR already handled all of them — exit loop
                break

            except ElementClickInterceptedException:
                print("Click was intercepted, retrying...")
                time.sleep(1)
                
        
        time.sleep(2)
        driver.find_element(By.XPATH, "//button[span[@title='Return Dashboard']]").click()
        time.sleep(2)

        month_str = row['month'].strip().title()
        month_num = month_name_to_number.get(month_str)
        if not month_num:
            driver.quit()
            continue
        quarter_num = get_quarter_from_month(month_num)
        dropdown_text = quarter_text_map.get(quarter_num)
        if len(month_str)==3:
            month_str=month_small_to_big.get(month_str)
        fill_timeline_details(driver, row['financial year'], dropdown_text, month_str)

        try:
            # --- 1. Find and Click GSTR2B Download Button ---
            gstr2b_xpath = (
                "//p[@data-ng-bind='x.return_ty' and text()='GSTR2B']"
                "/ancestor::div[contains(@data-ng-class, 'disableTable')]//button[text()='Download']"
            )
            gstr2b_download_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, gstr2b_xpath))
            )
            gstr2b_download_button.click()

            # --- 2. Find and Click "Generate Excel File" Button ---
            generate_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'GENERATE EXCEL FILE')]"))
            )

            # Get file list *before* clicking, in case it's a single (Case 1) download
            before_files = set(os.listdir(download_path))
            generate_button.click()

            # Wait for page to react (either start download or show table)
            time.sleep(5)

            downloaded_files_list = []  # Will store all downloaded file paths
            table_links_xpath = "//table[contains(@class, 'align-table')]//a[text()='Download here']"

            try:
                # --- 3. (Case 2) Check for the table of multiple files ---
                short_wait = WebDriverWait(driver, 5)  # 5-second wait for the table
                download_links = short_wait.until(
                    EC.presence_of_all_elements_located((By.XPATH, table_links_xpath))
                )

                print(f"[i] Found {len(download_links)} files in the table. Downloading all...")

                for i in range(len(download_links)):
                    try:
                        # Re-find link each time to avoid StaleElementReferenceException
                        current_link = short_wait.until(
                            EC.presence_of_all_elements_located((By.XPATH, table_links_xpath))
                        )[i]

                        link_before_files = set(os.listdir(download_path))  # Check files right before click
                        current_link.click()

                        new_filename = wait_for_new_file(download_path, link_before_files, timeout=TIMEOUT)

                        if new_filename:
                            full_path = os.path.abspath(os.path.join(download_path, new_filename))
                            downloaded_files_list.append(full_path)
                            print(f"  [✓] Downloaded: {new_filename}")
                            time.sleep(2)  # Brief pause
                        else:
                            print(f"  [✘] Clicked a link but no download was detected.")
                    except Exception as e:
                        print(f"  [✘] Error clicking download link #{i + 1}: {e}")

            except TimeoutException:
                # --- 3. (Case 1) No table appeared, check for single file ---
                print("[i] No download table found. Checking for single file download.")
                new_filename = wait_for_new_file(download_path, before_files, timeout=TIMEOUT)

                if new_filename:
                    full_path = os.path.abspath(os.path.join(download_path, new_filename))
                    downloaded_files_list.append(full_path)

            # --- 4. Process all downloaded files (for both cases) ---
            if downloaded_files_list:
                df.at[index, 'status'] = 'done'
                print(f"[✓] Total files downloaded for {row['username']}: {len(downloaded_files_list)}")

                final_attachment_path = ""

                if len(downloaded_files_list) == 1:
                    # If only one file, send it directly
                    final_attachment_path = downloaded_files_list[0]
                    df.at[index, 'downloaded_file_path'] = final_attachment_path
                else:
                    # If multiple files, ZIP them
                    zip_filename = (
                        f"{row['username']}_GSTR2B_{row['month']}_{row['financial year']}.zip"
                    ).replace("/", "-")
                    final_attachment_path = os.path.abspath(os.path.join(download_path, zip_filename))

                    print(f"[i] Zipping {len(downloaded_files_list)} files into {zip_filename}...")
                    with zipfile.ZipFile(final_attachment_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                        for file_path in downloaded_files_list:
                            zipf.write(file_path, os.path.basename(file_path))

                    df.at[index, 'downloaded_file_path'] = final_attachment_path  # Store the zip path

                # Send the email with the single attachment (either the 1 file or the 1 zip)
                send_email_with_attachment(
                    to_email=row['email'],
                    file_path=final_attachment_path,
                    sender_email=sender_email,
                    sender_password=sender_password,
                    email_subject=email_subject,
                    email_body=email_body
                )

            else:
                # No files were downloaded in EITHER case
                df.at[index, 'status'] = 'no'
                print(f"[✘] Download failed or timed out for user {row['username']}")

        except (NoSuchElementException, TimeoutException) as e:
            # This block runs if the GSTR2B *main download button* was not found
            print(f"[✘] Could not find the 'GSTR2B' Download button for user {row['username']}. Skipping. Error: {e}")
            df.at[index, 'status'] = 'no'
             
        driver.quit()
        completed += 1
        progress_callback(completed, total)
        
    df.to_excel(excel_path, index=False)