@echo off
setlocal EnableDelayedExpansion

set "ADDIN_NAME=Sarwar_PCMS"
set "ADDIN_FILE=Sarwar_PCMS.xll"
set "SOURCE_DIR=DLL Files"
set "SOURCE_XLL=%SOURCE_DIR%\Sarwar_PCMS.xll"
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
::  PRE-CHECK - Source File Exists?
:: ============================================================
if not exist "%SOURCE_XLL%" (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
        "Add-Type -AssemblyName PresentationFramework;" ^
        "[System.Windows.MessageBox]::Show('Source file not found. Please check the DLL Files folder.','Sarwar PCMS','OK','Error')" >nul 2>&1
    exit /b 1
)

:: ============================================================
::  STEP 1 - Create Folders (clean any stale leftover install first)
:: ============================================================
if exist "%INSTALL_DIR%" rd /s /q "%INSTALL_DIR%" >nul 2>&1
if not exist "%INSTALL_DIR%"        mkdir "%INSTALL_DIR%"       >nul 2>&1

:: ============================================================
::  STEP 2 - Copy XLL (Keep Original Name)
:: ============================================================
copy /y "%SOURCE_XLL%" "%INSTALL_DIR%\%ADDIN_FILE%" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
        "Add-Type -AssemblyName PresentationFramework;" ^
        "[System.Windows.MessageBox]::Show('Failed to copy XLL file. Check folder permissions.','Sarwar PCMS','OK','Error')" >nul 2>&1
    exit /b 1
)

:: ============================================================
::  STEP 3 - Copy Any Supporting Files (if present)
:: ============================================================
if exist "%SOURCE_DIR%\*.dll" (
    copy /y "%SOURCE_DIR%\*.dll" "%INSTALL_DIR%\" >nul 2>&1
)
if exist "%SOURCE_DIR%\*.dna" (
    copy /y "%SOURCE_DIR%\*.dna" "%INSTALL_DIR%\" >nul 2>&1
)
if exist "%SOURCE_DIR%\runtimes" (
    if not exist "%INSTALL_DIR%\runtimes" mkdir "%INSTALL_DIR%\runtimes" >nul 2>&1
    xcopy /y /e /i "%SOURCE_DIR%\runtimes" "%INSTALL_DIR%\runtimes" >nul 2>&1
)

:: ============================================================
::  STEP 4 - Unblock All Files
:: ============================================================
powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
    "Get-ChildItem -Path '%INSTALL_DIR%' -Recurse | Unblock-File -EA SilentlyContinue" >nul 2>&1

:: ============================================================
::  STEP 5 - Defender Exclusion (Admin Only - Silent Skip)
:: ============================================================
net session >nul 2>&1
if %ERRORLEVEL%==0 (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
        "Add-MpPreference -ExclusionPath '%INSTALL_DIR%' -EA SilentlyContinue" >nul 2>&1
)

:: ============================================================
::  STEP 6 - Registry
::  Part A : Trusted Location  (Fixes 1004 Error)
::  Part B : Smart OPEN Slot   (Registers XLL in Excel)
:: ============================================================
powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
    "$base='HKCU:\Software\Microsoft\Office\16.0\Excel';" ^
    "$optPath=\"$base\Options\";" ^
    "$trustPath=\"$base\Security\Trusted Locations\SarwarPCMS\";" ^
    "$xllTarget='/R \"%INSTALL_DIR%\Sarwar_PCMS.xll\"';" ^
    "if(-not (Test-Path $trustPath)){New-Item -Path $trustPath -Force|Out-Null};" ^
    "Set-ItemProperty -Path $trustPath -Name 'Path'            -Value '%INSTALL_DIR%\' -Force;" ^
    "Set-ItemProperty -Path $trustPath -Name 'AllowSubfolders' -Value 1 -PropertyType DWord -Force;" ^
    "Set-ItemProperty -Path $trustPath -Name 'Description'     -Value 'Sarwar PCMS' -Force;" ^
    "$props=Get-ItemProperty -Path $optPath -EA SilentlyContinue;" ^
    "$slots=$props.PSObject.Properties|Where-Object{$_.Name -match '^OPEN\d*$'};" ^
    "$existing=$slots|Where-Object{$_.Value -like '*Sarwar_PCMS*'};" ^
    "if($existing){" ^
    "  Set-ItemProperty -Path $optPath -Name $existing.Name -Value $xllTarget -Force" ^
    "} else {" ^
    "  $used=$slots|Select-Object -ExpandProperty Name;" ^
    "  $slot='OPEN';$i=1;" ^
    "  while($used -contains $slot){$slot='OPEN'+$i;$i++}" ^
    "  New-ItemProperty -Path $optPath -Name $slot -Value $xllTarget -PropertyType String -Force|Out-Null" ^
    "}" >nul 2>&1

:: ============================================================
::  DONE
:: ============================================================
powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command ^
    "Add-Type -AssemblyName PresentationFramework;" ^
    "[System.Windows.MessageBox]::Show('Installation complete. Please restart Excel.','Sarwar PCMS','OK','Information')" >nul 2>&1

exit /b
endlocal
