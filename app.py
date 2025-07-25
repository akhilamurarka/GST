import streamlit as st
import pandas as pd
import os, time, requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from pydub import AudioSegment
from pydub.utils import which
from word_extractor import transcribe_wav_to_digits
import pyautogui

# Set ffmpeg and ffprobe paths manually
AudioSegment.converter = which("ffmpeg")
AudioSegment.ffprobe = which("ffprobe")

# Month and quarter mapping
month_name_to_number = {
    "January": 1, "Jan": 1, "February": 2, "Feb": 2, "March": 3, "Mar": 3,
    "April": 4, "Apr": 4, "May": 5, "June": 6, "Jun": 6,
    "July": 7, "Jul": 7, "August": 8, "Aug": 8, "September": 9, "Sep": 9,
    "October": 10, "Oct": 10, "November": 11, "Nov": 11, "December": 12, "Dec": 12,
}
quarter_text_map = {
    "1": "Quarter 1 (Apr - Jun)", "2": "Quarter 2 (Jul - Sep)",
    "3": "Quarter 3 (Oct - Dec)", "4": "Quarter 4 (Jan - Mar)"
}

def get_quarter_from_month(month_num):
    if 1 <= month_num <= 3:
        return "4"
    elif 4 <= month_num <= 6:
        return "1"
    elif 7 <= month_num <= 9:
        return "2"
    elif 10 <= month_num <= 12:
        return "3"
    return None

def Chrome_setup(download_path):
    options = Options()
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_argument("--start-maximized")
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=options)

def Captcha_solving(driver, index, wait, df):
    for attempt in range(3):
        try:
            audio_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@ng-click='play()']")))
            audio_button.click()
            time.sleep(2)

            audio_element = driver.find_element(By.ID, "audioCap")
            audio_src = audio_element.get_attribute("src")
            driver.execute_script(f"window.open('{audio_src}', '_blank');")
            driver.switch_to.window(driver.window_handles[1])
            time.sleep(2)

            session = requests.Session()
            for cookie in driver.get_cookies():
                session.cookies.set(cookie["name"], cookie["value"])

            audio_mp3 = f"captcha_{index+1}.mp3"
            audio_wav = f"captcha_{index+1}.wav"
            response = session.get(audio_src)
            if "audio" not in response.headers.get("Content-Type", ""):
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                continue

            with open(audio_mp3, "wb") as f:
                f.write(response.content)

            driver.close()
            driver.switch_to.window(driver.window_handles[0])

            AudioSegment.from_mp3(audio_mp3).export(audio_wav, format="wav")
            digits = transcribe_wav_to_digits(audio_wav)
            if len(digits) == 6:
                driver.find_element(By.ID, "captcha").clear()
                driver.find_element(By.ID, "captcha").send_keys(digits)
                driver.find_element(By.CSS_SELECTOR, "button.btn.btn-primary").click()
                df.at[index, 'status'] = 'yes'
                return True
            else:
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
    Select(driver.find_element(By.NAME, "fin")).select_by_visible_text(year)
    time.sleep(1)
    Select(driver.find_element(By.NAME, "quarter")).select_by_visible_text(quarter_text)
    time.sleep(1)
    Select(driver.find_element(By.NAME, "mon")).select_by_visible_text(month)
    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR, "button.btn.btn-primary.srchbtn").click()
    time.sleep(2)

def run_automation(excel_path, download_path, progress_callback):
    df = pd.read_excel(excel_path)
    total = len(df)
    completed = 0

    for index, row in df.iterrows():
        if row['status'] != 'no':
            completed += 1
            continue

        driver = Chrome_setup(download_path)
        wait = WebDriverWait(driver, 10)
        driver.get("https://services.gst.gov.in/services/login")

        wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(row["username"])
        driver.find_element(By.ID, "user_pass").send_keys(row["password"])

        if not Captcha_solving(driver, index, wait, df):
            driver.quit()
            continue

        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@ng-click='cancelcallback()']"))).click()
        except TimeoutException:
            pass

        driver.find_element(By.XPATH, "//button[span[@title='Return Dashboard']]").click()
        time.sleep(2)

        month_str = row['month'].strip().title()
        month_num = month_name_to_number.get(month_str)
        if not month_num:
            driver.quit()
            continue
        quarter_num = get_quarter_from_month(month_num)
        dropdown_text = quarter_text_map.get(quarter_num)

        fill_timeline_details(driver, row['financial year'], dropdown_text, month_str)

        all_buttons = driver.find_elements(By.XPATH, "//button[text()='Download']")
        if len(all_buttons) > 1:
            all_buttons[1].click()
            driver.find_element(By.XPATH, "//button[contains(text(), 'GENERATE EXCEL FILE')]").click()
            time.sleep(3)
            pyautogui.press('enter')
            time.sleep(5)

        driver.quit()
        completed += 1
        progress_callback(completed, total)

    df.to_excel(excel_path, index=False)

# Streamlit UI
st.set_page_config(layout="centered")
st.title("📄 Automated GST Downloader")

excel_path = st.text_input("Excel file path", placeholder="e.g., C:/path/to/credentials.xlsx")
download_path = st.text_input("Download folder path", placeholder="e.g., C:/Downloads")

if st.button("Process"):
    if not os.path.exists(excel_path):
        st.error("❌ Excel file path not found.")
    elif not os.path.isdir(download_path):
        st.error("❌ Download folder path is invalid.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()

        def update_progress(completed, total):
            progress = completed / total
            progress_bar.progress(progress)
            status_text.text(f"✅ Downloaded {completed} out of {total} files.")

        run_automation(excel_path, download_path, update_progress)
        st.success("🎉 Automation complete!")
