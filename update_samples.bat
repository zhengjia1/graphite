@echo off
rem Refresh Graphite's sample datasets (samples\*.xlsx and the list in index.html)
rem from SingStat and data.gov.sg. Nothing is uploaded.
setlocal
cd /d "%~dp0"
chcp 65001 >nul

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH".
  goto :end
)

python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Installing openpyxl...
  python -m pip install --user --quiet openpyxl
  if errorlevel 1 (
    echo Couldn't install openpyxl. Try running:  python -m pip install --user openpyxl
    goto :end
  )
)

echo Downloading the latest data...
echo.
python tools\update_samples.py
if errorlevel 1 (
  echo.
  echo Finished with errors - see above.
) else (
  echo.
  echo Done.
)

:end
echo.
pause
