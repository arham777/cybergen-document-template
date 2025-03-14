@echo off
echo CyberGen Document Formatter POC Setup
echo =====================================
echo.

echo Installing required packages...
pip install python-docx

echo.
echo Creating template document...
python create_template.py

echo.
echo Preparing to run the POC...
echo You will be prompted to enter text or use the sample text from sample_test.txt
echo.
echo To use the sample text, copy and paste it when prompted, then type END on a new line.
echo.
pause

echo.
python cybergen_poc.py

echo.
echo POC demonstration complete!
echo.
pause 