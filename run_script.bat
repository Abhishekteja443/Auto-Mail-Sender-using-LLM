@echo off
call C:\Users\gabhi\Projects\Auto_mail_sender\venv\Scripts\activate.bat

python C:\Users\gabhi\Projects\Auto_mail_sender\resend.py >>C:\Users\gabhi\Projects\Auto_mail_sender\logfile.txt 2>&1

call C:\Users\gabhi\Projects\Auto_mail_sender\venv\Scripts\deactivate.bat

exit