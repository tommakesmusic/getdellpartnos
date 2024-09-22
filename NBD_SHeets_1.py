import keyring
from exchangelib import Account, Credentials
import re
import os

# Step 1: Authenticate and Connect to Outlook using Keyring
def connect_outlook(email):
    # Retrieve the password from the keyring for the provided email
    password = keyring.get_password('Outlook', email)
    if not password:
        raise Exception(f"No password found in keyring for email: {email}")

    credentials = Credentials(email, password)
    account = Account(email, credentials=credentials, autodiscover=True)
    return account

# Step 2: Download sys_report.pdf from "North NBDs" email
def find_and_download_email(account, subject, attachment_name):
    for item in account.inbox.filter(subject__icontains=subject):
        for attachment in item.attachments:
            if attachment.name == attachment_name:
                file_path = os.path.join("/path/to/save", attachment_name)  # Update with the correct path
                with open(file_path, 'wb') as f:
                    f.write(attachment.content)
                return file_path
    return None

# Step 3: Extract 11-digit numbers from "PUDO" or "Ni pudo" email
def extract_numbers_from_email(account, subject_keywords):
    for subject_keyword in subject_keywords:
        for item in account.inbox.filter(subject__icontains=subject_keyword):
            # Find 11-digit numbers in the email body
            numbers = re.findall(r'\b\d{11}\b', item.body)
            if numbers:
                return numbers
    return None

# Step 4: Search PDF for a specific number and extract the page number
def find_number_in_pdf(pdf_path, number):
    import PyPDF2
    with open(pdf_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        for page_num in range(len(reader.pages)):
            page_text = reader.pages[page_num].extract_text()
            if number in page_text:
                return page_num + 1  # Page numbers are 1-based
    return None

# Step 5: Extract and print the page containing the number
def extract_and_print_page(pdf_path, page_number):
    from PyPDF2 import PdfWriter, PdfReader
    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    writer.add_page(reader.pages[page_number - 1])

    output_pdf_path = f'/path/to/save/output_page_{page_number}.pdf'  # Update with the correct path
    with open(output_pdf_path, 'wb') as f:
        writer.write(f)

    # Simulating print by indicating which page is saved (replace with actual printing code if needed)
    print(f"Page {page_number} has been saved to {output_pdf_path}")

# Main Function
def main():
    email = "your-email@domain.com"  # Replace with your actual email

    # Connect to the Outlook account using Keyring
    outlook_account = connect_outlook(email)

    # Step 2: Download sys_report.pdf from "North NBDs" email
    pdf_path = find_and_download_email(outlook_account, "North NBDs", "sys_report.pdf")
    if not pdf_path:
        print("PDF not found")
        return

    # Step 3: Extract 11-digit numbers from "PUDO" or "Ni pudo" email
    subject_keywords = ["PUDO", "Ni pudo"]
    numbers = extract_numbers_from_email(outlook_account, subject_keywords)
    if not numbers:
        print("No 11-digit numbers found")
        return

    # Step 4 & 5: Search PDF for each number and print the page
    for number in numbers:
        page_number = find_number_in_pdf(pdf_path, number)
        if page_number:
            print(f"Number {number} found on page {page_number}, printing...")
            extract_and_print_page(pdf_path, page_number)
        else:
            print(f"Number {number} not found in the PDF.")

# Run the main function
if __name__ == "__main__":
    main()
