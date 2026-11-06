@echo off
rem Creates one pre-filled Power BI template per client from the workbooks in the Input folder.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Generate-ClientReports.ps1" -InputFolder "%~dp0Input" -TemplatePath "%~dp0Automated_Client_User_Billing_Report_Template.pbit" -OutputFolder "%~dp0Output"
echo.
pause
