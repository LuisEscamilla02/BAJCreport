import streamlit as st
import os
import tempfile
import json
from docx import Document
import mammoth
from combinedstaffandrepreports import (
    connect_to_google_sheets,
    get_sheet_names,
    process_staff_evaluation_docx
)

# Configuration file for persisting Spreadsheet ID
CONFIG_FILE = "config.json"

def load_config():
    try:
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f)

def load_data(credentials_path, spreadsheet_id):
    try:
        sheets_api = connect_to_google_sheets(credentials_path, spreadsheet_id)
        sheet_names = get_sheet_names(sheets_api, spreadsheet_id)
        return sheets_api, sheet_names
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None, []

def main():
    st.set_page_config(page_title="Staff Evaluation Report Generator", layout="centered")
    st.title("📄 Staff Evaluation Report Generator")
    st.markdown("""
        <style>
        html, body, [class*="css"]  {
            color: white !important;
            background-color: #111827 !important;
        }
        </style>
    """, unsafe_allow_html=True)

    # Load persisted config
    config = load_config()
    default_sheet_id = config.get("spreadsheet_id", "")

    with st.sidebar:
        st.header("🔧 Configuration")

        # Credentials uploader
        uploaded_file = st.file_uploader("📁 Upload Google API Credentials (.json)", type="json")
        if uploaded_file:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as tmp_file:
                tmp_file.write(uploaded_file.read())
                credentials_path = tmp_file.name
        else:
            credentials_path = None

        # Spreadsheet ID input (with persisted default)
        spreadsheet_id = st.text_input(
            "Google Spreadsheet ID",
            value=default_sheet_id,
            help="Enter the ID of the Google Sheet you want to use"
        )

        # Button to save the entered Spreadsheet ID
        if st.button("💾 Save Spreadsheet ID"):
            if spreadsheet_id.strip():
                config["spreadsheet_id"] = spreadsheet_id.strip()
                save_config(config)
                st.success("Spreadsheet ID saved for future use.")
            else:
                st.warning("Cannot save empty Spreadsheet ID.")

        # Load sheets button
        if st.button("🔄 Load Sheets"):
            if not credentials_path:
                st.warning("Please upload your Google API credentials JSON file.")
            elif not spreadsheet_id.strip():
                st.warning("Please enter a valid Google Spreadsheet ID.")
            else:
                sheets_api, sheet_names = load_data(credentials_path, spreadsheet_id)
                if sheet_names:
                    st.session_state.sheets_api = sheets_api
                    st.session_state.sheet_names = sheet_names
                    st.success(f"Loaded {len(sheet_names)} sheets.")
                else:
                    st.error("No sheets found or failed to load.")

    # Main area: generate report once sheets are loaded
    if "sheet_names" in st.session_state:
        selected_sheet = st.selectbox("📄 Select a Sheet", st.session_state.sheet_names)
        staff_name = st.text_input("👤 Staff Name to Process")

        if st.button("📤 Generate Report"):
            if not credentials_path:
                st.warning("Please upload your Google API credentials JSON file.")
            elif not staff_name.strip():
                st.warning("Please enter a staff name before generating the report.")
            else:
                try:
                    response = process_staff_evaluation_docx(
                        credentials_path,
                        spreadsheet_id,
                        selected_sheet,
                        staff_name.strip()
                    )
                    st.success(f"Report generated for {staff_name.strip()} ✅")

                    if staff_name.strip() in response:
                        file_path = response[staff_name.strip()]

                        # Download Button
                        with open(file_path, "rb") as f:
                            st.download_button(
                                label="📥 Download Report",
                                data=f,
                                file_name=os.path.basename(file_path),
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )

                        # Preview DOCX content using Mammoth
                        st.subheader("Report Preview")
                        with open(file_path, "rb") as docx_file:
                            result = mammoth.convert_to_html(docx_file)
                            html_preview = result.value

                        styled_preview = f"""
                        <div style="width: 100%; background-color: #111827; padding: 20px; overflow-y: auto; height: 800px;">
                            <div style="max-width: 900px; margin: auto; font-family: Arial, sans-serif; color: white; line-height: 1.6;">
                                {html_preview}
                            </div>
                        </div>
                        """
                        st.components.v1.html(styled_preview, height=850, scrolling=True)

                except Exception as e:
                    st.error(f"Failed to generate report: {e}")

if __name__ == "__main__":
    main()
