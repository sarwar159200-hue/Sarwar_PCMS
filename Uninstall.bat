@echo off
setlocal EnableDelayedExpansion

set "ADDIN_FILE=Sarwar_PCMS.xll"
set "INSTALL_DIR=%USERPROFILE%\Documents\Sarwar_PCMS"

:: ============================================================
::  PRE-CHECK - Excel Running?
:: ============================================================
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul
if %ERRORLEVEL%==0 (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
        "Add-Type -AssemblyName PresentationFramework;" ^
        "[System.Windows.MessageBox]::Show('Please close Excel first, then run this again.','Sarwar PCMS','OK','Warning')" >nul 2>&1
    exit /b 1
)

:: ============================================================
::  PRE-CHECK - Is it Installed?
:: ============================================================
if not exist "%INSTALL_DIR%" (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
        "Add-Type -AssemblyName PresentationFramework;" ^
        "[System.Windows.MessageBox]::Show('Sarwar PCMS is not installed.','Sarwar PCMS','OK','Information')" >nul 2>&1
    exit /b 0
)

:: ============================================================
::  STEP 1 - Remove Registry OPEN Slot
:: ============================================================
powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
    "$regPath='HKCU:\Software\Microsoft\Office\16.0\Excel\Options';" ^
    "$props=Get-ItemProperty -Path $regPath -EA SilentlyContinue;" ^
    "$slot=$props.PSObject.Properties|Where-Object{$_.Name -match '^OPEN\d*$' -and $_.Value -like '*Sarwar_PCMS*'};" ^
    "if($slot){Remove-ItemProperty -Path $regPath -Name $slot.Name -Force -EA SilentlyContinue}" >nul 2>&1

:: ============================================================
::  STEP 2 - Remove Trusted Location Registry Key
:: ============================================================
powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
    "$trustPath='HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\SarwarPCMS';" ^
    "if(Test-Path $trustPath){Remove-Item -Path $trustPath -Force -EA SilentlyContinue}" >nul 2>&1

:: ============================================================
::  STEP 3 - Remove Defender Exclusion (Admin Only)
:: ============================================================
net session >nul 2>&1
if %ERRORLEVEL%==0 (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
        "Remove-MpPreference -ExclusionPath '%INSTALL_DIR%' -EA SilentlyContinue" >nul 2>&1
)

:: ============================================================
::  STEP 4 - Delete Installation Folder (Everything)
:: ============================================================
rd /s /q "%INSTALL_DIR%" >nul 2>&1

:: ============================================================
::  DONE
:: ============================================================
if not exist "%INSTALL_DIR%" (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
        "Add-Type -AssemblyName PresentationFramework;" ^
        "[System.Windows.MessageBox]::Show('Uninstall complete. Please restart Excel.','Sarwar PCMS','OK','Information')" >nul 2>&1
) else (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
        "Add-Type -AssemblyName PresentationFramework;" ^
        "[System.Windows.MessageBox]::Show('Could not remove folder. Please delete manually: %INSTALL_DIR%','Sarwar PCMS','OK','Warning')" >nul 2>&1
)
exit /b
endlocal
