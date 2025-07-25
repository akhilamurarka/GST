import pandas as pd
import os, requests, time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from pydub import AudioSegment
from selenium.common.exceptions import TimeoutException
from pydub.utils import which
from word_extractor import transcribe_wav_to_digits
import pyautogui

# Set ffmpeg and ffprobe paths manually if needed
AudioSegment.converter = which("ffmpeg")
AudioSegment.ffprobe = which("ffprobe")

excel_file_path = "data/credentials.xlsx"

try:
    # Load credentials from Excel
    df = pd.read_excel(excel_file_path)  # Columns: username, password, financial_year, month, status
    if 'downloaded_file_path' not in df.columns:
        df['downloaded_file_path'] = ''  # Add empty column
except Exception as e:
    print(f"[ERROR] Could not load Excel file: {e}")
    exit(1)

# Month name to number map
month_name_to_number = {
    "January": 1, "Jan": 1,
    "February": 2, "Feb": 2,
    "March": 3, "Mar": 3,
    "April": 4, "Apr": 4,
    "May": 5, "May": 5,
    "June": 6, "Jun": 6,
    "July": 7, "Jul": 7,
    "August": 8, "Aug": 8,
    "September": 9, "Sep": 9,
    "October": 10, "Oct": 10,
    "November": 11, "Nov": 11,
    "December": 12, "Dec": 12,
}

def get_quarter_from_month(month_num):
    if 1 <= month_num <= 3:
        return "4"  # Jan-Mar => Q4 (of financial year)
    elif 4 <= month_num <= 6:
        return "1"  # Apr-Jun => Q1
    elif 7 <= month_num <= 9:
        return "2"  # Jul-Sep => Q2
    elif 10 <= month_num <= 12:
        return "3"  # Oct-Dec => Q3
    else:
        return None

quarter_text_map = {
    "1": "Quarter 1 (Apr - Jun)",
    "2": "Quarter 2 (Jul - Sep)",
    "3": "Quarter 3 (Oct - Dec)",
    "4": "Quarter 4 (Jan - Mar)"
}

def Chrome_setup():
    try:
        options = Options()
        prefs = {
            "download.prompt_for_download": True,  # disables the Save As dialog
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        options.add_argument("--start-maximized")
        options.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(options=options)
        return driver
    except Exception as e:
        print(f"[ERROR] Could not initialize Chrome driver: {e}")
        raise

def Captcha_solving(driver, index):
    MAX_RETRIES = 3
    success = False

    for attempt in range(1, MAX_RETRIES + 1):
        print(f"\n🔁 Attempt {attempt} for CAPTCHA solving...")

        try:
            wait = WebDriverWait(driver, 10)

            # Step 3: Click Audio Button
            audio_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@ng-click='play()']")))
            audio_button.click()
            time.sleep(2)

            # Step 4: Get Audio URL
            audio_element = driver.find_element(By.ID, "audioCap")
            time.sleep(1)
            audio_src = audio_element.get_attribute("src")
            time.sleep(2)

            # Step 5: Open Audio URL in New Tab
            driver.execute_script(f"window.open('{audio_src}', '_blank');")
            time.sleep(1)
            driver.switch_to.window(driver.window_handles[1])
            time.sleep(2)

            # Step 6: Download MP3 using Session Cookies
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

            # Step 7: Close Audio Tab & Return to Login
            driver.close()
            time.sleep(1)
            driver.switch_to.window(driver.window_handles[0])

            # Step 8: Convert MP3 → WAV
            try:
                AudioSegment.from_mp3(audio_mp3).export(audio_wav, format="wav")
            except Exception as e:
                print(f"❌ Audio conversion failed: {e}")
                continue

            # Step 9: Transcribe and Predict Digits
            digits = transcribe_wav_to_digits(audio_wav)
            print(f"[🔎 Predicted Digits]: {digits}")

            if len(digits) == 6:
                # Step 10: Fill CAPTCHA & Submit
                driver.find_element(By.ID, "captcha").clear()
                time.sleep(1)
                driver.find_element(By.ID, "captcha").send_keys(digits)
                time.sleep(1)
                driver.find_element(By.CSS_SELECTOR, "button.btn.btn-primary").click()
                time.sleep(1)
                success = True
                df.at[index, 'status'] = 'yes'
                return success
            else:
                print("❌ Could not extract 6 digits from audio. Retrying...")
                wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@ng-click='refreshCaptcha()']"))).click()

        except Exception as e:
            print(f"⚠️ Exception in attempt {attempt}: {e}")

        finally:
            # Clean up for next attempt
            if os.path.exists(audio_mp3):
                os.remove(audio_mp3)
            if os.path.exists(audio_wav):
                os.remove(audio_wav)
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

    if not success:
        print(f"[⛔] Failed to solve CAPTCHA for user after {MAX_RETRIES} attempts.")
        driver.quit()
        return success

def fill_timeline_details(driver, financial_year, dropdown_text, month_str):
    try:
        # Locate the <select> element
        year_dropdown = Select(driver.find_element(By.NAME, "fin"))
        time.sleep(1)
        # Select the year by visible text
        year_dropdown.select_by_visible_text(financial_year)
        time.sleep(1)

        # Locate the <select> element
        quarter_dropdown = Select(driver.find_element(By.NAME, "quarter"))
        time.sleep(1)
        # Select the quarter by visible text
        quarter_dropdown.select_by_visible_text(dropdown_text)
        time.sleep(1)

        # Locate the <select> element
        month_dropdown = Select(driver.find_element(By.NAME, "mon"))
        time.sleep(1)
        # Select the month by visible text
        month_dropdown.select_by_visible_text(month_str)
        time.sleep(1)

        driver.find_element(By.CSS_SELECTOR, "button.btn.btn-primary.srchbtn").click()
        time.sleep(2)
    except Exception as e:
        print(f"[ERROR] Failed to fill timeline details: {e}")
        raise

for index, row in df.iterrows():
    try:
        print(f"\n🔁 Processing user {index+1}: {row['username']}")

        financial_year = row['financial year']
        month_str = row['month'].strip().title()  # from Excel row, like "April"
        month_num = month_name_to_number.get(month_str)
        if month_num:
            quarter_num = get_quarter_from_month(month_num)
            dropdown_text = quarter_text_map.get(quarter_num)
        else:
            print(f"Invalid month name in Excel: {month_str}")
            df.at[index, 'status'] = 'invalid month'
            continue

        status = row.get('status', 'no')

        if status == "no":
            driver = None
            try:
                # Setup Chrome
                driver = Chrome_setup()
                wait = WebDriverWait(driver, 10)

                # Step 1: Open Login Page
                base_url = "https://services.gst.gov.in"
                driver.get(f"{base_url}/services/login")

                # Step 2: Fill Username & Password
                wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(row["username"])
                time.sleep(2)
                driver.find_element(By.ID, "user_pass").send_keys(row["password"])
                time.sleep(2)

                success = Captcha_solving(driver, index)
                if not success:
                    print(f"[✘] CAPTCHA solving failed for user {row['username']}. Skipping...")
                    df.at[index, 'status'] = 'captcha failed'
                    continue

                # Step 11: Wait & Save Result Page
                time.sleep(3)
                try:
                    remind_button = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//a[@ng-click='cancelcallback()']"))
                    )
                    remind_button.click()
                    print("Remind button clicked.")
                except TimeoutException:
                    print("Remind button not present, continuing...")

                time.sleep(2)
                driver.find_element(By.XPATH, "//button[span[@title='Return Dashboard']]").click()
                time.sleep(2)

                fill_timeline_details(driver, financial_year, dropdown_text, month_str)

                all_buttons = driver.find_elements(By.XPATH, "//button[text()='Download']")
                time.sleep(1)
                if len(all_buttons) > 1:
                    all_buttons[1].click()
                    time.sleep(1)
                else:
                    print("[!] Download button not found.")
                    df.at[index, 'status'] = 'download button missing'
                    continue

                driver.find_element(By.XPATH, "//button[contains(text(), 'GENERATE EXCEL FILE')]").click()

                time.sleep(3)
                pyautogui.press('enter')
                time.sleep(7)

                # Estimate or construct downloaded file path (based on browser behavior)
                download_folder = os.path.expanduser("~/Downloads")  # or replace with custom path if known
                expected_filename = f"gst_data_{row['username']}.xlsx"  # adjust this pattern if the downloaded file has a different name
                downloaded_file_path = os.path.join(download_folder, expected_filename)

                print(f"[✔] User {row['username']} completed. File saved to: {downloaded_file_path}")
                df.at[index, 'status'] = 'done'
                df.at[index, 'downloaded_file_path'] = downloaded_file_path

            except Exception as e_inner:
                print(f"[ERROR] Exception during automation for user {row['username']}: {e_inner}")
                df.at[index, 'status'] = 'error'
            finally:
                if driver:
                    driver.quit()

        else:
            print(f"Skipping user {row['username']} with status: {status}")

    except Exception as e_outer:
        print(f"[FATAL] Exception processing user index {index}: {e_outer}")
        df.at[index, 'status'] = 'fatal error'

# Save the updated Excel with status updates
try:
    df.to_excel(excel_file_path, index=False)
    print("\n📁 Updated Excel file saved with new statuses.")
except Exception as e_save:
    print(f"[ERROR] Could not save Excel file: {e_save}")
