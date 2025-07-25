# 🧾 GST Portal Automation with Selenium & Streamlit

This project automates the login and data download process from the [GST Portal](https://services.gst.gov.in) for multiple users listed in an Excel sheet.

It includes:
- Automated login via Selenium
- CAPTCHA solving (audio-based using speech recognition)
- Automated navigation to the return dashboard
- File download handling (Excel or PDF returns)
- Streamlit interface to run everything easily
- Logs and updates download status + file path in Excel
---

## 📂 Folder Structure

GST-Automation/
├── app.py                  # Main Streamlit + Automation Script
├── credentials.xlsx        # Input Excel file with user credentials
├── word_extractor.py       # Custom module to transcribe audio CAPTCHA
├── README.md               # This file

---

## 📌 Features

- ✅ Fully automated GST login for multiple users
- 🔐 Handles CAPTCHA via audio-to-text conversion
- 📥 Downloads return files and simulates "Save As" using `pyautogui`
- 📊 Streamlit interface to choose Excel input file and download location
- 📁 Automatically tracks downloaded file paths and updates status

---

## 📄 Input Excel Format

Your Excel file (`credentials.xlsx`) must include the following columns:

| username  | password | financial year | month   | status |
|-----------|----------|----------------|---------|--------|
| user1     | pass123  | 2023-24        | April   | no     |
| user2     | pass456  | 2023-24        | June    | no     |

- `status` should be set to `no` to process.
- After processing:
  - If successful: updated to `done`
  - If failed: updated to `fatal error` or remains `no`

---

## 🚀 How to Run

### ⚙️ Step 1: Install Requirements

It's recommended to use a virtual environment.

```bash
python -m venv venv
venv\Scripts\activate     # On Windows
pip install -r requirements.txt
```

If you don’t have a requirements.txt, install dependencies manually:
```
pip install selenium streamlit pandas openpyxl pydub pyautogui
```

Also install `ffmpeg` (required for audio CAPTCHA decoding):

- Download from https://ffmpeg.org/download.html
- Extract it and add the folder to your system's PATH
- Ensure `ffmpeg.exe` is accessible in command line



### ▶️ Step 2: Run the App
```
streamlit run app.py
```

- This will open a browser-based interface
- You’ll be asked to enter the Excel file path and a download folder path

---

## 🧠 Tech Stack

- Python 3.10+
- Selenium for browser automation
- Streamlit for UI
- pydub + ffmpeg for audio CAPTCHA decoding
- pyautogui to simulate file saving
- pandas + openpyxl for Excel file management

---

## ⚠️ Notes

- Make sure your Chrome version matches your ChromeDriver version.
- Run the script with permissions to allow file downloads.
- CAPTCHA decoding relies on audio quality and may occasionally fail.
- Streamlit and Chrome must both run on the same environment.
- `Save As` is simulated with `pyautogui`, so don’t move the mouse while it's running.

---

## 🛠️ Future Enhancements

- Headless Chrome support
- Retry mechanism via UI
- Visual progress tracker (e.g., 3 of 10 completed)
- Download detection and auto-renaming

---

## 🧾 Output

After successful execution:
- The Excel file will be updated with:
  - `status` of each user
  - `downloaded_file_path` column (auto-added if missing)

---

## 📬 Contact

For bugs or suggestions, please open an issue or reach out to the maintainer.

---
