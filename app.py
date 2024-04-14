import os
import re
import PyPDF2
import textract
import streamlit as st
import xlwt

def extract_info_from_pdf(pdf_path):
    text = ''
    email = ''
    contact = ''

    # Open PDF file
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        num_pages = len(pdf_reader.pages)

        # Extract text from each page
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()

    # Extract email using regex
    email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    email_match = re.search(email_regex, text)
    if email_match:
        email = email_match.group(0)

    # Extract contact number using regex
    contact_regex = r'\b\d{10}\b'
    contact_match = re.search(contact_regex, text)
    if contact_match:
        contact = contact_match.group(0)

    return email, contact, text

def extract_info_from_doc(doc_path):
    text = textract.process(doc_path).decode('utf-8')
    email = ''
    contact = ''

    # Extract email using regex
    email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    email_match = re.search(email_regex, text)
    if email_match:
        email = email_match.group(0)

    # Extract contact number using regex
    contact_regex = r'\b\d{10}\b'
    contact_match = re.search(contact_regex, text)
    if contact_match:
        contact = contact_match.group(0)

    return email, contact, text

def extract_info_from_cv(cv_path):
    _, ext = os.path.splitext(cv_path)
    if ext == '.pdf':
        return extract_info_from_pdf(cv_path)
    elif ext == '.docx':
        return extract_info_from_doc(cv_path)
    else:
        print(f"Unsupported file format: {ext}")
        return '', '', ''

def save_to_excel(data, output_path):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('CV Data')

    # Headers
    sheet.write(0, 0, 'Email')
    sheet.write(0, 1, 'Contact')
    sheet.write(0, 2, 'Text')

    # Data
    for row, (email, contact, text) in enumerate(data, start=1):
        sheet.write(row, 0, email)
        sheet.write(row, 1, contact)
        sheet.write(row, 2, text)

    workbook.save(output_path)

def main():
    st.title("CV Parser")

    uploaded_files = st.file_uploader("Upload CVs", accept_multiple_files=True, type=['pdf', 'docx'])

    if st.button("Parse CVs"):
        data = []

        # Iterate over uploaded files
        for uploaded_file in uploaded_files:
            cv_content = uploaded_file.getvalue()
            cv_path = f"./temp/{uploaded_file.name}"
            with open(cv_path, 'wb') as f:
                f.write(cv_content)

            email, contact, text = extract_info_from_cv(cv_path)
            data.append((email, contact, text))

        # Save extracted data to Excel
        save_to_excel(data, 'cv_data.xls')
        st.success("CVs parsed successfully. Download the Excel file below.")
        st.download_button(label="Download CV Data", data=open("cv_data.xls", "rb").read(), file_name="cv_data.xls", mime="application/vnd.ms-excel")

if __name__ == "__main__":
    main()