@echo off
setlocal

echo ================================
echo Building Conveyor Calculator
echo ================================

set VBC="C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe"

%VBC% ^
 /target:winexe ^
 /out:ConveyorCalculator.exe ^
 /reference:System.dll ^
 /reference:System.Windows.Forms.dll ^
 /reference:System.Drawing.dll ^
 ConveyorCalculatorForm.vb ^
 Program.vb

if errorlevel 1 (
    echo BUILD FAILED
    pause
    exit /b
)

echo BUILD SUCCESS
echo Output: ConveyorCalculator.exe
pause
