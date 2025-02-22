@echo off
call venv\Scripts\activate.bat

python resend.py >> logfile.txt 2>&1

call venv\Scripts\deactivate.bat

exit
