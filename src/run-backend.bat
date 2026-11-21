@echo off
rem Start the Office Editor backend (office-backend) on http://127.0.0.1:47831.
rem Double-click this file, or run:  run-backend.bat [--port PORT]
rem Keep the window open while you use the extension; close it to stop the backend.
setlocal
cd /d "%~dp0"
title office-backend

where cargo >nul 2>nul
if %ERRORLEVEL%==0 (
  echo Building and starting office-backend ^(the first build takes a few minutes^)...
  cargo run --release -p office-backend -- %*
  goto :end
)

if exist "target\release\office-backend.exe" (
  echo cargo not found - starting the previously built target\release\office-backend.exe
  "target\release\office-backend.exe" %*
  goto :end
)

echo.
echo Rust is not installed, so office-backend cannot be built.
echo  1. Install Rust from https://rustup.rs  (accept the defaults)
echo  2. Open a NEW terminal and run this script again.
echo.
echo The extension still works without the backend for .docx, .xlsx, .xls, .ods,
echo .csv, .tsv, .md and .txt files; .doc, .rtf and .odt need the backend.

:end
echo.
echo office-backend stopped.
pause
endlocal
