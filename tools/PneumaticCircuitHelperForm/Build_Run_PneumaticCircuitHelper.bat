@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul
title MetaMech Pneumatic Circuit Helper - Build and Run

REM --- Always run from this BAT folder ---
pushd "%~dp0"

set "WORKDIR=%CD%"
set "SRC=%WORKDIR%\PneumaticCircuitHelperForm.vb"
set "BOOT=%WORKDIR%\Program_PneumaticCircuitHelper.vb"
set "OUT=%WORKDIR%\PneumaticCircuitHelper.exe"
set "LOG=%WORKDIR%\build_log_pneumatic_circuit_helper.txt"
set "VBC=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\vbc.exe"

if not exist "%VBC%" set "VBC=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\vbc.exe"

echo MetaMech Pneumatic Circuit Helper - Build and Run
echo WorkingDir: %WORKDIR%
echo Log: %LOG%
echo.

(
echo ============================================================
echo MetaMech Pneumatic Circuit Helper - BUILD LOG
echo %date%  %time%
echo ============================================================
echo Compiler: %VBC%
echo WorkingDir: %WORKDIR%
echo SRC: %SRC%
echo BOOT: %BOOT%
echo OUT: %OUT%
) > "%LOG%"

if not exist "%VBC%" (
  echo [ERROR] Compiler not found
  >> "%LOG%" echo [ERROR] Compiler not found
  pause
  exit /b 1
) else (
  echo [OK] Found compiler
  >> "%LOG%" echo [OK] Found compiler
)

if not exist "%SRC%" (
  echo [ERROR] Source file not found: %SRC%
  >> "%LOG%" echo [ERROR] Source file not found: %SRC%
  pause
  exit /b 1
) else (
  echo [OK] Found source
  >> "%LOG%" echo [OK] Found source
)

if not exist "%BOOT%" (
  echo [ERROR] Program file not found: %BOOT%
  >> "%LOG%" echo [ERROR] Program file not found: %BOOT%
  pause
  exit /b 1
) else (
  echo [OK] Found Program file
  >> "%LOG%" echo [OK] Found Program file
)

echo Compiling...
>> "%LOG%" echo Compiling...

REM /target:winexe = no console window for final app
REM /main forces correct startup module
"%VBC%" ^
 /nologo ^
 /target:winexe ^
 /main:Program_PneumaticCircuitHelper ^
 /optionstrict+ ^
 /optionexplicit+ ^
 /optioninfer+ ^
 /imports:System,System.Drawing,System.Windows.Forms,System.IO,System.Text ^
 /reference:"System.dll" ^
 /reference:"System.Drawing.dll" ^
 /reference:"System.Windows.Forms.dll" ^
 /out:"%OUT%" ^
 "%BOOT%" "%SRC%" >> "%LOG%" 2>&1

set "ERR=%ERRORLEVEL%"
>> "%LOG%" echo ExitCode: %ERR%

if not "%ERR%"=="0" (
  echo.
  echo [ERROR] Build failed. See log:
  echo %LOG%
  type "%LOG%"
  pause
  exit /b %ERR%
)

echo.
echo [SUCCESS] Build OK
echo EXE: %OUT%
echo Running tool...
echo.
>> "%LOG%" echo [SUCCESS] Build OK

REM Run and keep window open after it exits
start "" "%OUT%"

echo If the app does not appear, check:
echo   %WORKDIR%\PneumaticCircuitHelper_startup_log.txt
echo.
pause
exit /b 0