@echo off
setlocal

set "APP_URL=http://127.0.0.1/projeto-aniversario/app/"
set "APACHE_EXE=C:\xampp\apache\bin\httpd.exe"

tasklist /FI "IMAGENAME eq httpd.exe" | find /I "httpd.exe" >nul
if errorlevel 1 (
  if exist "%APACHE_EXE%" (
    start "" "%APACHE_EXE%"
    timeout /t 2 /nobreak >nul
  )
)

if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" (
  start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --app="%APP_URL%"
  exit /b 0
)

if exist "C:\Program Files\Microsoft\Edge\Application\msedge.exe" (
  start "" "C:\Program Files\Microsoft\Edge\Application\msedge.exe" --app="%APP_URL%"
  exit /b 0
)

if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
  start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --app="%APP_URL%"
  exit /b 0
)

if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
  start "" "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --app="%APP_URL%"
  exit /b 0
)

start "" "%APP_URL%"
