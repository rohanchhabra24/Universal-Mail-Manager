# Universal Mail Manager

A sleek, self-contained Python email client built with Streamlit. It uses native Python libraries (`imaplib` & `smtplib`) to read, parse, and send emails across major providers like Outlook and Gmail. 

## Features
- **Universal Login**: Safely log in using an App Password — no Azure AD or Google Cloud Console configurations needed!
- **Rich UI**: Filter emails, identify flags/attachments, and view HTML emails accurately through an immersive Streamlit popup.
- **Compose & Send**: Fully functional composition form supporting 'To', 'CC', Attachments and High Priority markings. 

## Installation

1. Clone the repository and install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the Streamlit app:
   ```bash
   streamlit run outlook_mail_manager.py
   ```

## Security Note
This app asks for your provider Email and App Password. Your credentials are fundamentally stored only within the active Streamlit session state memory and are entirely wiped when Streamlit restarts or you click Log Out. 

## Technologies
- Streamlit
- IMAP / SMTP (Python Standard Library)
