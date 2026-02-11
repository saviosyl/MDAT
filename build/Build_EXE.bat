@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

title MetaMech MDAT - EXE Builder

REM ============================================================
REM  PROJECT ROOT (Build_EXE.bat is inside MDAT\build\)
REM ============================================================
set "ROOT=%~dp0.."

pushd "%ROOT%" >nul 2>&1
if errorlevel 1 (
  echo ============================================================
  echo [ERROR] Cannot access project root:
  echo   %ROOT%
  echo ============================================================
  echo.
  pause
  exit /b 1
)

cls
echo ============================================================
echo   MetaMech Mechanical Design Automation - EXE Builder
echo ============================================================
echo Root: %CD%
echo Time: %DATE% %TIME%
echo.

REM ------------------------------------------------------------
REM PATHS
REM ------------------------------------------------------------
set "APP_DIR=app"
set "TOOLS_DIR=app\tools"
set "OUTPUT_DIR=output"

set "OUT_EXE=%OUTPUT_DIR%\MDAT.exe"
set "LIC_EXE=%OUTPUT_DIR%\LicenseGenerator.exe"

set "LOG_MDAT=%OUTPUT_DIR%\build_mdat.log"
set "LOG_LIC=%OUTPUT_DIR%\build_licensegen.log"

set "VBC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\vbc.exe"

REM ------------------------------------------------------------
REM PDF MERGE TOOL (ADDED)
REM ------------------------------------------------------------
set "PDF_MERGE_SRC=tools\PdfMergeTool.vb"
set "PDF_DLL=%OUTPUT_DIR%\PdfSharp-gdi.dll"
set "PDF_MERGE_EXE=%OUTPUT_DIR%\PdfMergeTool.exe"
set "LOG_PDFMERGE=%OUTPUT_DIR%\build_pdfmerge.log"

REM ------------------------------------------------------------
REM CHECK PROJECT FILES
REM ------------------------------------------------------------
echo ============================================================
echo [CHECK] Project files
echo ============================================================
if not exist "%APP_DIR%\Program.vb" (
  echo [ERROR] Missing: %APP_DIR%\Program.vb
  echo.
  pause
  popd
  exit /b 1
)
if not exist "%APP_DIR%\MainForm.vb" (
  echo [ERROR] Missing: %APP_DIR%\MainForm.vb
  echo.
  pause
  popd
  exit /b 1
)
echo [OK] app files found
echo.

REM ------------------------------------------------------------
REM CHECK COMPILER
REM ------------------------------------------------------------
echo ============================================================
echo [CHECK] VB.NET Compiler
echo ============================================================
if not exist "%VBC%" (
  echo [ERROR] vbc.exe not found:
  echo   %VBC%
  echo.
  pause
  popd
  exit /b 1
)
echo [OK] Found: %VBC%
echo.

REM ------------------------------------------------------------
REM PREP OUTPUT
REM ------------------------------------------------------------
echo ============================================================
echo [STEP] Cleaning previous build
echo ============================================================
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%" >nul 2>&1
del /f /q "%OUT_EXE%" "%LIC_EXE%" "%PDF_MERGE_EXE%" >nul 2>&1
del /f /q "%LOG_MDAT%" "%LOG_LIC%" "%LOG_PDFMERGE%" >nul 2>&1
echo [OK] Output folder ready: %OUTPUT_DIR%
echo.

REM ------------------------------------------------------------
REM REQUIRE SEAT FILES (in app\)
REM ------------------------------------------------------------
echo ============================================================
echo [CHECK] Seat enforcement files
echo ============================================================
if not exist "%APP_DIR%\SeatEnforcer.vb" (
  echo [ERROR] Missing: %APP_DIR%\SeatEnforcer.vb
  echo         Put this file inside: %APP_DIR%\
  echo.
  pause
  popd
  exit /b 1
)
if not exist "%APP_DIR%\SeatServerClient.vb" (
  echo [ERROR] Missing: %APP_DIR%\SeatServerClient.vb
  echo         Put this file inside: %APP_DIR%\
  echo.
  pause
  popd
  exit /b 1
)
if not exist "%APP_DIR%\SeatCache.vb" (
  echo [ERROR] Missing: %APP_DIR%\SeatCache.vb
  echo         Put this file inside: %APP_DIR%\
  echo.
  pause
  popd
  exit /b 1
)
if not exist "%APP_DIR%\MachineFingerprint.vb" (
  echo [ERROR] Missing: %APP_DIR%\MachineFingerprint.vb
  echo         Put this file inside: %APP_DIR%\
  echo.
  pause
  popd
  exit /b 1
)
echo [OK] Seat files found
echo.

REM ------------------------------------------------------------
REM PHASE 1 TELEMETRY FILES (ADDED - REQUIRED)
REM ------------------------------------------------------------
echo ============================================================
echo [CHECK] Telemetry (Phase 1) files
echo ============================================================
if not exist "%APP_DIR%\TelemetryClient.vb" (
  echo [ERROR] Missing: %APP_DIR%\TelemetryClient.vb
  echo         Put this file inside: %APP_DIR%\
  echo.
  pause
  popd
  exit /b 1
)
if not exist "%APP_DIR%\TelemetryJson.vb" (
  echo [ERROR] Missing: %APP_DIR%\TelemetryJson.vb
  echo         Put this file inside: %APP_DIR%\
  echo.
  pause
  popd
  exit /b 1
)
if not exist "%APP_DIR%\TelemetryQueue.vb" (
  echo [ERROR] Missing: %APP_DIR%\TelemetryQueue.vb
  echo         Put this file inside: %APP_DIR%\
  echo.
  pause
  popd
  exit /b 1
)
if not exist "%APP_DIR%\TelemetryService.vb" (
  echo [ERROR] Missing: %APP_DIR%\TelemetryService.vb
  echo         Put this file inside: %APP_DIR%\
  echo.
  pause
  popd
  exit /b 1
)
echo [OK] Telemetry files found
echo.

REM ============================================================
REM BUILD MAIN APPLICATION
REM ============================================================
echo ============================================================
echo [STEP] COMPILING MAIN APPLICATION  (MDAT.exe)
echo ============================================================
echo Log: %LOG_MDAT%
echo.

"%VBC%" ^
 /nologo ^
 /target:winexe ^
 /platform:x86 ^
 /optimize+ ^
 /optionstrict+ ^
 /optionexplicit+ ^
 /main:Program ^
 /reference:"System.dll" ^
 /reference:"System.Core.dll" ^
 /reference:"System.Drawing.dll" ^
 /reference:"System.Windows.Forms.dll" ^
 /reference:"System.Security.dll" ^
 /out:"%OUT_EXE%" ^
 "%APP_DIR%\Program.vb" ^
 "%APP_DIR%\MainForm.vb" ^
 "%APP_DIR%\SeatEnforcer.vb" ^
 "%APP_DIR%\SeatServerClient.vb" ^
 "%APP_DIR%\SeatCache.vb" ^
 "%APP_DIR%\MachineFingerprint.vb" ^
 "%APP_DIR%\TelemetryClient.vb" ^
 "%APP_DIR%\TelemetryJson.vb" ^
 "%APP_DIR%\TelemetryQueue.vb" ^
 "%APP_DIR%\TelemetryService.vb" ^
 "%APP_DIR%\AboutForm.vb" ^
 "%APP_DIR%\UITheme.vb" ^
 "%TOOLS_DIR%\ThemeApplier.vb" ^
 "%TOOLS_DIR%\UISettings.vb" ^
 "%TOOLS_DIR%\AdminLoginForm.vb" ^
 "%TOOLS_DIR%\AdminMacroEditorForm.vb" ^
 "%TOOLS_DIR%\ConveyorCalculatorForm.vb" ^
 "%TOOLS_DIR%\ConveyorModeForm.vb" ^
 "%TOOLS_DIR%\Engineering Calculator.vb" ^
 "%TOOLS_DIR%\EngineeringNotepadForm.vb" ^
 "%TOOLS_DIR%\PneumaticCylinderCalculatorForm.vb" ^
 "%TOOLS_DIR%\PneumaticCylinderToolForm.vb" ^
 "%TOOLS_DIR%\AirConsumptionForm.vb" ^
 "%TOOLS_DIR%\LicenseInfo.vb" ^
 "%TOOLS_DIR%\Licensing.vb" ^
 "%TOOLS_DIR%\MachineId.vb" ^
 "%TOOLS_DIR%\MotorPowerForm.vb" ^
 "%TOOLS_DIR%\PdfExportForm.vb" ^
 "%TOOLS_DIR%\PdfQuoteForm.vb" ^
 "%TOOLS_DIR%\ProjectManager.vb" ^
 "%TOOLS_DIR%\ResultContext.vb" ^
 "%TOOLS_DIR%\SeatPool.vb" ^
 "%TOOLS_DIR%\SolidWorksGate.vb" ^
 "%TOOLS_DIR%\TierLocks.vb" ^
 "%TOOLS_DIR%\TorqueCalculatorForm.vb" ^
 "%TOOLS_DIR%\UnitConverterForm.vb" ^
 "%TOOLS_DIR%\UnitSystem.vb" ^
 1>"%LOG_MDAT%" 2>&1

if errorlevel 1 goto :MDAT_FAIL

REM Extra safety: sometimes build returns success but exe is missing
if not exist "%OUT_EXE%" goto :MDAT_MISSING

echo.
echo ============================================================
echo [SUCCESS] MAIN APPLICATION BUILD OK
echo ============================================================
echo Output: %OUT_EXE%
echo.

REM ============================================================
REM BUILD LICENSE GENERATOR (WINFORMS UI)
REM ============================================================
echo ============================================================
echo [STEP] COMPILING LICENSE GENERATOR  (LicenseGenerator.exe)
echo ============================================================
echo Log: %LOG_LIC%
echo.

if not exist "%TOOLS_DIR%\LicenseGeneratorProgram.vb" (
  echo [ERROR] Missing: %TOOLS_DIR%\LicenseGeneratorProgram.vb
  echo.
  pause
  popd
  exit /b 1
)
if not exist "%TOOLS_DIR%\LicenseGeneratorForm.vb" (
  echo [ERROR] Missing: %TOOLS_DIR%\LicenseGeneratorForm.vb
  echo.
  pause
  popd
  exit /b 1
)
if not exist "%TOOLS_DIR%\LicenseGeneratorCore.vb" (
  echo [ERROR] Missing: %TOOLS_DIR%\LicenseGeneratorCore.vb
  echo.
  pause
  popd
  exit /b 1
)

"%VBC%" ^
 /nologo ^
 /target:winexe ^
 /platform:anycpu ^
 /optimize+ ^
 /optionstrict+ ^
 /optionexplicit+ ^
 /main:LicenseGeneratorProgram ^
 /reference:"System.dll" ^
 /reference:"System.Core.dll" ^
 /reference:"System.Drawing.dll" ^
 /reference:"System.Windows.Forms.dll" ^
 /reference:"System.Security.dll" ^
 /out:"%LIC_EXE%" ^
 "%TOOLS_DIR%\LicenseGeneratorProgram.vb" ^
 "%TOOLS_DIR%\LicenseGeneratorForm.vb" ^
 "%TOOLS_DIR%\LicenseGeneratorCore.vb" ^
 1>"%LOG_LIC%" 2>&1

if errorlevel 1 goto :LIC_FAIL

if not exist "%LIC_EXE%" goto :LIC_MISSING

echo.
echo ============================================================
echo [SUCCESS] LICENSE GENERATOR BUILD OK
echo ============================================================
echo Output: %LIC_EXE%
echo.

REM ============================================================
REM BUILD PDF MERGE TOOL (ADDED)  (PdfMergeTool.exe)
REM ============================================================
echo ============================================================
echo [STEP] COMPILING PDF MERGE TOOL  (PdfMergeTool.exe)
echo ============================================================
echo Log: %LOG_PDFMERGE%
echo.

if not exist "%PDF_MERGE_SRC%" (
  echo [ERROR] Missing: %PDF_MERGE_SRC%
  echo         Put this file here: MDAT\tools\PdfMergeTool.vb
  echo.
  pause
  popd
  exit /b 1
)

if not exist "%PDF_DLL%" (
  echo [ERROR] Missing: %PDF_DLL%
  echo         Put this DLL here: MDAT\output\PdfSharp-gdi.dll
  echo.
  pause
  popd
  exit /b 1
)

"%VBC%" ^
 /nologo ^
 /target:exe ^
 /platform:x86 ^
 /optimize+ ^
 /optionstrict+ ^
 /optionexplicit+ ^
 /reference:"%PDF_DLL%" ^
 /out:"%PDF_MERGE_EXE%" ^
 "%PDF_MERGE_SRC%" ^
 1>"%LOG_PDFMERGE%" 2>&1

if errorlevel 1 goto :PDFMERGE_FAIL

if not exist "%PDF_MERGE_EXE%" goto :PDFMERGE_MISSING

echo.
echo ============================================================
echo [SUCCESS] PDF MERGE TOOL BUILD OK
echo ============================================================
echo Output: %PDF_MERGE_EXE%
echo.

REM ============================================================
REM FINAL SUMMARY
REM ============================================================
echo ============================================================
echo [SUCCESS] BUILD COMPLETE - ALL OK
echo ============================================================
echo Created:
echo   - %OUT_EXE%
echo   - %LIC_EXE%
echo   - %PDF_MERGE_EXE%
echo Logs:
echo   - %LOG_MDAT%
echo   - %LOG_LIC%
echo   - %LOG_PDFMERGE%
echo ============================================================
echo.
pause
popd
endlocal
exit /b 0

REM ============================================================
REM FAIL HANDLERS
REM ============================================================
:MDAT_FAIL
echo.
echo ============================================================
echo [FAILED] MAIN APPLICATION BUILD FAILED
echo ============================================================
echo Log saved to: %LOG_MDAT%
echo.
echo ---- LAST 80 LINES (MDAT LOG) ----
powershell -NoProfile -Command "if (Test-Path '%LOG_MDAT%') { Get-Content '%LOG_MDAT%' -Tail 80 } else { 'Log file not found.' }"
echo ---------------------------------
echo.
pause
popd
endlocal
exit /b 1

:MDAT_MISSING
echo.
echo ============================================================
echo [FAILED] MDAT.exe NOT CREATED
echo ============================================================
echo Build command finished but output EXE is missing.
echo Log saved to: %LOG_MDAT%
echo.
echo ---- LAST 80 LINES (MDAT LOG) ----
powershell -NoProfile -Command "if (Test-Path '%LOG_MDAT%') { Get-Content '%LOG_MDAT%' -Tail 80 } else { 'Log file not found.' }"
echo ---------------------------------
echo.
pause
popd
endlocal
exit /b 1

:LIC_FAIL
echo.
echo ============================================================
echo [FAILED] LICENSE GENERATOR BUILD FAILED
echo ============================================================
echo Log saved to: %LOG_LIC%
echo.
echo ---- LAST 80 LINES (LICENSEGEN LOG) ----
powershell -NoProfile -Command "if (Test-Path '%LOG_LIC%') { Get-Content '%LOG_LIC%' -Tail 80 } else { 'Log file not found.' }"
echo ----------------------------------------
echo.
pause
popd
endlocal
exit /b 1

:LIC_MISSING
echo.
echo ============================================================
echo [FAILED] LicenseGenerator.exe NOT CREATED
echo ============================================================
echo Build command finished but output EXE is missing.
echo Log saved to: %LOG_LIC%
echo.
echo ---- LAST 80 LINES (LICENSEGEN LOG) ----
powershell -NoProfile -Command "if (Test-Path '%LOG_LIC%') { Get-Content '%LOG_LIC%' -Tail 80 } else { 'Log file not found.' }"
echo ----------------------------------------
echo.
pause
popd
endlocal
exit /b 1

:PDFMERGE_FAIL
echo.
echo ============================================================
echo [FAILED] PDF MERGE TOOL BUILD FAILED
echo ============================================================
echo Log saved to: %LOG_PDFMERGE%
echo.
echo ---- LAST 80 LINES (PDFMERGE LOG) ----
powershell -NoProfile -Command "if (Test-Path '%LOG_PDFMERGE%') { Get-Content '%LOG_PDFMERGE%' -Tail 80 } else { 'Log file not found.' }"
echo -------------------------------------
echo.
pause
popd
endlocal
exit /b 1

:PDFMERGE_MISSING
echo.
echo ============================================================
echo [FAILED] PdfMergeTool.exe NOT CREATED
echo ============================================================
echo Build command finished but output EXE is missing.
echo Log saved to: %LOG_PDFMERGE%
echo.
echo ---- LAST 80 LINES (PDFMERGE LOG) ----
powershell -NoProfile -Command "if (Test-Path '%LOG_PDFMERGE%') { Get-Content '%LOG_PDFMERGE%' -Tail 80 } else { 'Log file not found.' }"
echo -------------------------------------
echo.
pause
popd
endlocal
exit /b 1
