@echo off
setlocal
title DocShift Setup
set "APP=%LOCALAPPDATA%\DocShift"
python -c "pass" >nul 2>nul
if errorlevel 1 (
  echo Python is needed. Installing it now with winget...
  winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
  echo.
  echo When Python finishes installing, close this window and run Install-Windows.bat again.
  pause
  exit /b 1
)
echo [1/4] Copying DocShift...
xcopy "%~dp0*" "%APP%\" /E /I /Y /Q >nul
echo [2/4] Installing components (needs internet, one time only, about a minute)...
python -m venv "%APP%\venv"
"%APP%\venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
"%APP%\venv\Scripts\python.exe" -m pip install --quiet -r "%APP%\requirements.txt"
if errorlevel 1 (
  echo.
  echo Could not download the components. Connect to the internet and run this again.
  pause
  exit /b 1
)
echo Installing the higher-quality PDF to Word engine (optional)...
"%APP%\venv\Scripts\python.exe" -m pip install --quiet pdf2docx
echo [3/4] Checking for LibreOffice (needed for Word, Excel and PowerPoint conversions)...
if not exist "%ProgramFiles%\LibreOffice\program\soffice.exe" if not exist "%ProgramFiles(x86)%\LibreOffice\program\soffice.exe" (
  choice /m "LibreOffice was not found. Install it now (free, about 350 MB)"
  if not errorlevel 2 winget install -e --id TheDocumentFoundation.LibreOffice --accept-package-agreements --accept-source-agreements
)
echo [4/4] Creating shortcuts...
powershell -NoProfile -ExecutionPolicy Bypass -File "%APP%\make-shortcut.ps1" -App "%APP%" -Pyw "%APP%\venv\Scripts\pythonw.exe"
echo.
echo Done. A DocShift icon is now on your desktop and in the Start menu.
start "" "%APP%\venv\Scripts\pythonw.exe" "%APP%\app.py"
pause
