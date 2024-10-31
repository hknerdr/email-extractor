import os
import re
import pandas as pd
import streamlit as st
from io import BytesIO
from docx import Document
import PyPDF2

def extract_emails_from_files(files):
    email_set = set()
    email_pattern = re.compile(
        r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-z]{2,}', re.IGNORECASE)

    total_files = len(files)
    progress_bar = st.progress(0)
    log_area = st.empty()
    logs = []

    for idx, uploaded_file in enumerate(files):
        filename = uploaded_file.name
        file_extension = os.path.splitext(filename)[1].lower()
        logs.append(f"Processing file {idx + 1}/{total_files}: **{filename}**")
        log_area.markdown("\n".join(logs))
        try:
            if file_extension in ['.xls', '.xlsx', '.xlsm']:
                xls = pd.ExcelFile(uploaded_file)
                for sheet_name in xls.sheet_names:
                    logs.append(f"Reading sheet: {sheet_name}")
                    log_area.markdown("\n".join(logs))
                    df = pd.read_excel(
                        xls, sheet_name=sheet_name, dtype=str)
                    for column in df.columns:
                        for cell in df[column].dropna():
                            matches = email_pattern.findall(str(cell))
                            for email in matches:
                                email_set.add(email)
            elif file_extension == '.docx':
                doc = Document(uploaded_file)
                logs.append(f"Extracting text from Word document")
                log_area.markdown("\n".join(logs))
                for para in doc.paragraphs:
                    matches = email_pattern.findall(para.text)
                    for email in matches:
                        email_set.add(email)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            matches = email_pattern.findall(cell.text)
                            for email in matches:
                                email_set.add(email)
            elif file_extension == '.pdf':
                reader = PyPDF2.PdfReader(uploaded_file)
                logs.append(f"Extracting text from PDF")
                log_area.markdown("\n".join(logs))
                for page_num in range(len(reader.pages)):
                    page = reader.pages[page_num]
                    text = page.extract_text()
                    if text:
                        matches = email_pattern.findall(text)
                        for email in matches:
                            email_set.add(email)
            else:
                logs.append(f"⚠️ Unsupported file type: {filename}")
                log_area.markdown("\n".join(logs))
        except Exception as e:
            logs.append(f"❌ Error processing file {filename}: {e}")
            log_area.markdown("\n".join(logs))
        progress_bar.progress((idx + 1) / total_files)

    return sorted(email_set), logs

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Email Addresses')
    return output.getvalue()

def main():
    st.set_page_config(page_title="Email Address Extractor",
                       page_icon="📧", layout="wide")

    st.title("📧 Email Address Extractor from Various File Types")
    st.write("""
    Upload multiple files (`.xls`, `.xlsx`, `.xlsm`, `.docx`, `.pdf`) containing email addresses in various formats.
    The app will extract all unique email addresses and provide a consolidated Excel file for download.
    """)

    uploaded_files = st.file_uploader(
        "Choose Files",
        type=['xls', 'xlsx', 'xlsm', 'docx', 'pdf'],
        accept_multiple_files=True
    )

    if st.button("Extract Emails"):
        if uploaded_files:
            with st.spinner('Processing...'):
                emails, logs = extract_emails_from_files(uploaded_files)
            if emails:
                st.success(
                    f"Extraction complete! {len(emails)} unique email(s) found.")
                # Display the extracted emails in a table
                email_df = pd.DataFrame({'Email Addresses': emails})
                st.dataframe(email_df)

                # Add copy to clipboard functionality
                st.markdown("### Copy Emails to Clipboard")
                emails_str = '\n'.join(emails)
                st.code(emails_str)
                st.button("Copy to Clipboard", on_click=lambda: st.experimental_set_query_params(
                    emails=emails_str))

                # Convert DataFrame to Excel
                excel_data = convert_df_to_excel(email_df)
                st.download_button(
                    label="📥 Download Extracted Emails as Excel",
                    data=excel_data,
                    file_name='Extracted_Email_Addresses.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.warning("No email addresses found in the uploaded files.")
        else:
            st.warning("Please upload at least one file.")

    # Add a clear/reset button
    if st.button("Clear"):
        st.experimental_rerun()

if __name__ == "__main__":
    main()
