@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

title MetaMech - Package Trial ZIP

REM ============================================================
REM  Creates a clean trial ZIP ready to upload to website
REM  Run this AFTER Build_EXE.bat succeeds
REM ============================================================

set "ROOT=%~dp0.."
pushd "%ROOT%" >nul 2>&1

echo ============================================================
echo   MetaMech Trial Packager
echo ============================================================
echo Root: %CD%
echo Time: %DATE% %TIME%
echo.

set "OUTPUT=output"
set "STAGING=output\_package_staging"
set "ZIPNAME=MetaMech-Trial.zip"
set "ZIPPATH=output\%ZIPNAME%"

REM ------------------------------------------------------------
REM CHECK BUILD EXISTS
REM ------------------------------------------------------------
echo [CHECK] Build files...
if not exist "%OUTPUT%\MDAT.exe" (
  echo [ERROR] MDAT.exe not found. Run Build_EXE.bat first!
  pause
  popd
  exit /b 1
)
if not exist "%OUTPUT%\PdfMergeTool.exe" (
  echo [WARNING] PdfMergeTool.exe not found - continuing without it
)
echo [OK] Build files found
echo.

REM ------------------------------------------------------------
REM CLEAN STAGING
REM ------------------------------------------------------------
echo [STEP] Preparing staging folder...
if exist "%STAGING%" rmdir /s /q "%STAGING%"
mkdir "%STAGING%" >nul 2>&1
mkdir "%STAGING%\assets\logo" >nul 2>&1
mkdir "%STAGING%\assets\icons" >nul 2>&1
mkdir "%STAGING%\assets\themes" >nul 2>&1
echo [OK] Staging ready
echo.

REM ------------------------------------------------------------
REM COPY FILES (only what the customer needs)
REM ------------------------------------------------------------
echo [STEP] Copying files...

REM Main EXE
copy /y "%OUTPUT%\MDAT.exe" "%STAGING%\" >nul

REM PDF tools
if exist "%OUTPUT%\PdfMergeTool.exe" copy /y "%OUTPUT%\PdfMergeTool.exe" "%STAGING%\" >nul
if exist "%OUTPUT%\PdfSharp-gdi.dll" copy /y "%OUTPUT%\PdfSharp-gdi.dll" "%STAGING%\" >nul

REM Config (with server URL and client token)
copy /y "%OUTPUT%\Config.txt" "%STAGING%\" >nul

REM Public key (for offline license verification)
copy /y "%OUTPUT%\MetaMech_RSA_PUBLIC.xml" "%STAGING%\" >nul

REM Assets
if exist "%OUTPUT%\assets\logo\app.ico" copy /y "%OUTPUT%\assets\logo\app.ico" "%STAGING%\assets\logo\" >nul
if exist "%OUTPUT%\assets\logo\logo.png" copy /y "%OUTPUT%\assets\logo\logo.png" "%STAGING%\assets\logo\" >nul

if exist "%OUTPUT%\assets\icons\*.svg" copy /y "%OUTPUT%\assets\icons\*.svg" "%STAGING%\assets\icons\" >nul

if exist "%OUTPUT%\assets\themes\*.theme" copy /y "%OUTPUT%\assets\themes\*.theme" "%STAGING%\assets\themes\" >nul

REM UI settings defaults
if exist "%OUTPUT%\ui.ini" copy /y "%OUTPUT%\ui.ini" "%STAGING%\" >nul

echo [OK] Files copied
echo.

REM ------------------------------------------------------------
REM VERIFY: NO PRIVATE FILES INCLUDED
REM ------------------------------------------------------------
echo [CHECK] Security check...
set "FAIL=0"

if exist "%STAGING%\license.key" (
  echo [ERROR] license.key found in package! Removing...
  del /f /q "%STAGING%\license.key"
)
if exist "%STAGING%\MetaMech_RSA_PRIVATE.xml" (
  echo [ERROR] PRIVATE KEY found in package! Removing...
  del /f /q "%STAGING%\MetaMech_RSA_PRIVATE.xml"
  set "FAIL=1"
)
if exist "%STAGING%\Generate-License.ps1" (
  echo [ERROR] License generator script found! Removing...
  del /f /q "%STAGING%\Generate-License.ps1"
)
if exist "%STAGING%\LicenseGenerator.exe" (
  echo [ERROR] LicenseGenerator.exe found! Removing...
  del /f /q "%STAGING%\LicenseGenerator.exe"
)
if exist "%STAGING%\*.log" (
  del /f /q "%STAGING%\*.log" >nul 2>&1
)
if exist "%STAGING%\*.csv" (
  del /f /q "%STAGING%\*.csv" >nul 2>&1
)
if exist "%STAGING%\*.jsonl" (
  del /f /q "%STAGING%\*.jsonl" >nul 2>&1
)
if exist "%STAGING%\lic_debug.txt" (
  del /f /q "%STAGING%\lic_debug.txt"
)
if exist "%STAGING%\SeatConfig.txt" (
  del /f /q "%STAGING%\SeatConfig.txt"
)
if exist "%STAGING%\test_macro_server.cmd" (
  del /f /q "%STAGING%\test_macro_server.cmd"
)
if exist "%STAGING%\ui.settings" (
  del /f /q "%STAGING%\ui.settings"
)

if "%FAIL%"=="1" (
  echo [ERROR] Private key was almost included! Check your output folder.
  pause
  popd
  exit /b 1
)

echo [OK] No private files in package
echo.

REM ------------------------------------------------------------
REM CREATE ZIP
REM ------------------------------------------------------------
echo [STEP] Creating ZIP...

if exist "%ZIPPATH%" del /f /q "%ZIPPATH%"

REM Use PowerShell to create ZIP
powershell -NoProfile -Command ^
  "Compress-Archive -Path '%STAGING%\*' -DestinationPath '%ZIPPATH%' -Force"

if not exist "%ZIPPATH%" (
  echo [ERROR] ZIP creation failed!
  pause
  popd
  exit /b 1
)

echo [OK] ZIP created: %ZIPPATH%
echo.

REM ------------------------------------------------------------
REM CLEANUP STAGING
REM ------------------------------------------------------------
rmdir /s /q "%STAGING%" >nul 2>&1

REM ------------------------------------------------------------
REM SUMMARY
REM ------------------------------------------------------------
echo ============================================================
echo [SUCCESS] TRIAL PACKAGE READY
echo ============================================================
echo.
echo   Output:  %ZIPPATH%
echo.
echo   Contents:
echo     - MDAT.exe (main application)
echo     - PdfMergeTool.exe (PDF merge tool)
echo     - PdfSharp-gdi.dll (PDF library)
echo     - Config.txt (server config)
echo     - MetaMech_RSA_PUBLIC.xml (license verification)
echo     - assets\ (icons, themes, logo)
echo.
echo   NOT included (safe):
echo     - NO license.key (trial auto-activates from server)
echo     - NO private key
echo     - NO license generator
echo     - NO logs or debug files
echo.
echo   Upload this ZIP to metamechsolutions.com/download
echo ============================================================
echo.
pause
popd
endlocal
exit /b 0
