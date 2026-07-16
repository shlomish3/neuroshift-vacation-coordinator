@echo off
echo Deploying to Google Apps Script...
echo ----------------------------------------

call clasp push
call clasp deploy -i AKfycbwxOS-T-BJC2TsxwBMpXN1k6OHbh5b4ip0BuzUOERIfcuI-eC9PADIRzP6zjCVNV-EUdg

if %errorlevel% neq 0 (
    echo.
    echo Error deploying.
    echo Please make sure you have done the one-time setup:
    echo 1. Run "clasp login"
    echo 2. Run "clasp clone <ScriptID>"
    echo.
) else (
    echo.
    echo Deployment successful! Your Apps Script is now up to date.
)

pause
