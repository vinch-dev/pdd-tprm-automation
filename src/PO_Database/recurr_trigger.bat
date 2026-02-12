@echo off

cd /d "C:\01. PO_Hist_Database\"

echo -------------------------------------------------------
echo Work Hist Rate Report starting...
echo -------------------------------------------------------
pause
"python.exe" "PO_wk_hist-data_recurr_ver.0.0.1.py"
echo -------------------------------------------------------
echo Database update...
echo -------------------------------------------------------
pause
"python.exe" "PO_conv_db.py"
echo -------------------------------------------------------
echo Checking what's wrong...
echo -------------------------------------------------------
pause
