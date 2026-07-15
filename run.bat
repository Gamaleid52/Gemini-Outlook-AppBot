@echo off
echo ==========================================
echo  Checking and Installing Requirements...
echo ==========================================
pip install -r requirements.txt

echo.
echo ==========================================
echo  Running AI Email Automation Script...
echo ==========================================
py main.py

echo.
echo ==========================================
echo  Process Finished!
echo ==========================================
pause