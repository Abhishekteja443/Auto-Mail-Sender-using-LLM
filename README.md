# Auto Email Sender for Job Applications

## Overview
This project automates the process of sending and resending job application emails using OpenAI's Llama 3.2 model. It consists of a Streamlit-based application for generating and sending emails (`app.py`) and a scheduled script (`resend.py`) that checks logs and resends follow-up emails weekly.

## Features
- Automatically generates personalized job application emails.
- Validates email addresses before sending.
- Attaches resumes to the emails.
- Maintains a log (`logs.csv`) to track sent emails.
- Resends follow-up emails weekly if no response is received.

## Prerequisites
Ensure the following dependencies are installed before running the application:

- Python 3.8+
- Streamlit
- OpenAI's Ollama library
- `win32com.client` for Outlook email automation
- `csv` for logging

### Installation
1. Clone the repository:
   ```sh
   git clone https://github.com/yourusername/auto-mail-sender.git
   cd auto-mail-sender
   ```
2. Create a virtual environment and activate it:
   ```sh
   python -m venv venv
   venv\Scripts\activate  # Windows
   source venv/bin/activate  # Mac/Linux
   ```
3. Install dependencies:
   ```sh
   pip install -r requirements.txt
   ```

## Usage
### Running the Web App
1. Start the Streamlit application:
   ```sh
   streamlit run app.py
   ```
2. Enter the professor's name, email, and job description.
3. Click **Generate Email** to create an email using AI.
4. Click **Send Email** to send the email through Outlook.
5. The sent email details are stored in `logs.csv`.

### Running the Auto Resend Script
This script checks `logs.csv` and automatically resends follow-up emails every week.
```sh
python resend.py
```

## Configuration
- Update `logs.csv` with previous email details.
- Modify email templates and prompts in `app.py` and `resend.py` to suit your needs.
- Ensure Outlook is installed and configured for email automation.
- Change the `about_me` section in `app.py` and `resend.py` to reflect your personal details.
- Update the paths for your resume and `logs.csv` in the script to match your local directory structure.
- Update `logs.csv` with previous email details.
- Modify email templates and prompts in `app.py` and `resend.py` to suit your needs.
- Ensure Outlook is installed and configured for email automation.

## Notes
- The script only resends follow-up emails if the last sent email is older than a week.
- All log files (`app.log`, `resend.log`) help track errors and activities.

## License
This project is licensed under the MIT License.

## Author
Abhishek Teja Goli

