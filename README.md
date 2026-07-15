# AI-Powered Email Automation Script 🚀

A Python automation script that reads company data from an Excel sheet, leverages Google Gemini API to generate tailored cover letters, and automatically creates and saves formatted drafts in Microsoft Outlook with a PDF resume attached.

## Features
- **Excel Parsing:** Extracts company info dynamically.
- **AI Integration:** Uses Gemini 2.5 Flash for contextual generation.
- **Outlook Automation:** Automates draft creation & PDF attachments securely via `pywin32`.
- **Environment Isolation:** Keeps sensitive APIs and config secure via `.env`.