@echo off
rem Launch the CDR default-properties float window without a console.
setlocal

set "PYW=D:\Python 3.12.4\pythonw.exe"
if not exist "%PYW%" (
  for /f "delims=" %%i in ('where pythonw.exe 2^>nul') do (
    set "PYW=%%i"
    goto :found
  )
)
:found

if not exist "%PYW%" (
  echo [ERROR] pythonw.exe not found.
  echo Edit this file and set PYW to your pythonw.exe path.
  pause
  exit /b 1
)

start "" "%PYW%" "%~dp0cdr_defaults.pyw"
exit /b 0
