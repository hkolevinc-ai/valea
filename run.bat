@echo off
setlocal
cd /d "%~dp0"
python scraper.py
if errorlevel 1 (
  echo.
  echo Scraper failed. Review the error above.
  pause
  exit /b 1
)
echo.
echo Ready. Open the results folder.
pause
