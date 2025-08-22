import streamlit as st
import os
from backend import *

# Streamlit UI
st.set_page_config(layout="centered")
st.title("📄 Automated GST Downloader")

excel_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])
download_path = st.text_input("Download folder path", value=os.path.expanduser("~\\Downloads"))

with st.sidebar:
    sender_email = st.text_input("Sender's Email")
    sender_app_password = st.text_input("Sender's App Password", type="password")
    # Add the app password guide link here
    st.markdown(
        """
        🔐 Need help generating an App Password?  
        [Click here for Google’s guide](https://support.google.com/accounts/answer/185833)
        NOTE: Kindly remove the spaces present between the password
        """,
        unsafe_allow_html=True
    )
    email_subject = st.text_input("Email Subject", value="Your GST Excel File")
    email_body = st.text_area("Email Content", value="Please find your GST Excel file attached.")

if st.button("Process") and excel_file and download_path:
    if not os.path.isdir(download_path):
        st.error("❌ Download folder path is invalid.")
    elif not sender_email or not sender_app_password:
        st.error("❌ Please provide sender's email and app password in the sidebar.")
    else:
        try:
            # 🔄 Save uploaded file to a temp file
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                tmp_file.write(excel_file.read())
                tmp_excel_path = tmp_file.name

            progress_bar = st.progress(0)
            status_text = st.empty()

            def update_progress(completed, total):
                progress = completed / total
                progress_bar.progress(progress)
                status_text.text(f"✅ Downloaded {completed} out of {total} files.")

            # 🔄 Run automation on saved temp file
            run_automation(tmp_excel_path, download_path, update_progress,sender_email=sender_email,
                sender_password=sender_app_password,
                email_subject=email_subject,
                email_body=email_body)

            # 🔄 Load updated Excel file to memory for download
            with open(tmp_excel_path, "rb") as f:
                updated_excel_bytes = f.read()

            st.success("🎉 Automation complete! Download the updated Excel file below:")
            st.download_button(
                label="📥 Download Updated Excel",
                data=updated_excel_bytes,
                file_name="updated_credentials.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ An error occurred: {e}")

