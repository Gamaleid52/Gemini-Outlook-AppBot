import os
import time
import pandas as pd
from google import genai  # Official updated google-genai package
from dotenv import load_dotenv
import win32com.client as win32

# 1. Load configuration from .env file
load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
PDF_PATH = os.getenv("PDF_PATH")
BASE_PROMPT = os.getenv("BASE_PROMPT")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
EXCEL_PATH = os.getenv("EXCEL_PATH")

# Initialize the Gemini Client
client = genai.Client(api_key=GEMINI_API_KEY)


def generate_cover_letter(company_name, notes):
    """
    Connects to Gemini API using the official google-genai SDK to generate a tailored Cover Letter.
    """
    try:
        prompt = BASE_PROMPT.format(company_name=company_name, notes=notes)
        
        # Calling the updated Gemini 2.5 Flash model
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
        )
        return response.text.strip()
    except Exception as e:
        print(f"❌ Error generating text for {company_name}: {e}")
        return None


def create_outlook_draft(recipient_email, subject, body, attachment_path, sender_email):
    """
    Automates local MS Outlook via COM interface to create a Draft, attach PDF, and select sender account.
    """
    try:
        # Initialize COM connection to Outlook
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)  # 0 represents olMailItem
        
        # Try to set the specific "From" account if configured
        if sender_email:
            account_found = False
            for account in outlook.Session.Accounts:
                if account.SmtpAddress.lower() == sender_email.lower():
                    # Set the sending account using OLE binding
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                    account_found = True
                    break
            if not account_found:
                print(f"⚠️ Warning: Sender email {sender_email} not configured in Outlook. Using default account.")

        # Fill email details
        mail.To = recipient_email
        mail.Subject = subject
        mail.Body = body
        
        # Attach the PDF Resume
        if attachment_path and os.path.exists(attachment_path):
            mail.Attachments.Add(attachment_path)
            print(f"📎 PDF Resume attached successfully.")
        else:
            print(f"⚠️ Warning: PDF file not found at: {attachment_path}")

        # Save to Drafts
        mail.Save()
        return True
    except Exception as e:
        print(f"❌ Outlook COM Error: {e}")
        return False


def main():
    print("==========================================")
    print("   🚀 Starting AI Job Application Agent   ")
    print("==========================================")

    # 2. Validate and load Excel database
    if not EXCEL_PATH or not os.path.exists(EXCEL_PATH):
        print(f"❌ Error: Excel file not found at path: {EXCEL_PATH}")
        return
        
    try:
        df = pd.read_excel(EXCEL_PATH)
        print(f"📊 Excel file loaded. Found ({len(df)}) companies to process.")
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return

    # 3. Iterate through each company row
    for index, row in df.iterrows():
        company_name = row['CompanyName']
        recipient_email = row['Email']
        company_notes = row['Notes']
        
        print(f"\n------------------------------------------")
        print(f"⏳ Processing: {company_name} | {recipient_email}")
        print(f"------------------------------------------")
        
        # Step A: Generate tailored cover letter
        cover_letter = generate_cover_letter(company_name, company_notes)
        
        if cover_letter:
            print("✅ Cover letter successfully generated via Gemini AI.")
            subject = f"Application for Backend / Full-Stack Developer Position - {company_name}"
            
            # Step B: Interact with Outlook COM and save draft
            success = create_outlook_draft(
                recipient_email=recipient_email,
                subject=subject,
                body=cover_letter,
                attachment_path=PDF_PATH,
                sender_email=SENDER_EMAIL
            )
            
            if success:
                print(f"💾 Email successfully saved as Draft for {company_name}!")
        
        # Delay to respect local resources and Gemini API rate limits
        print("⏳ Waiting 6 seconds before next company...")
        time.sleep(6)

    print("\n==========================================")
    print("🎉 All companies processed successfully!")
    print("==========================================")


if __name__ == "__main__":
    main()