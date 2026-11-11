@echo off
REM ============================================================
REM  Builds PPTtoWord-Setup.exe  (run on Windows)
REM  Needs: Python 3.10+  and  Inno Setup 6 (https://jrsoftware.org/isdl.php)
REM ============================================================
setlocal
cd /d "%~dp0"

if not exist .venv ( python -m venv .venv )
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller || goto :error

REM 1) Package the app (folder mode = faster start-up, fewer antivirus false alarms)
pyinstaller --noconfirm --clean --windowed --name PPTtoWord ^
  --icon app.ico --add-data "app.ico;." ^
  --collect-data pptx --collect-data docx ^
  app.py || goto :error

REM 2) Wrap it into a proper installer
set ISCC="%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not exist %ISCC% (
  echo Inno Setup 6 not found. Install it from https://jrsoftware.org/isdl.php and run again.
  goto :error
)
%ISCC% installer.iss || goto :error

echo.
echo Done. Installer: installer_output\PPTtoWord-Setup.exe
exit /b 0

:error
echo.
echo Build failed - see messages above.
exit /b 1
