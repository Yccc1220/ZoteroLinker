@echo off
setlocal

:: Zotero Linker VSTO installer builder
:: Usage: build.bat [version] [config]
::   config: Debug or Release (defaults to Release)

set VERSION=%1
if "%VERSION%"=="" set VERSION=1.0.0

set CONFIG=%2
if "%CONFIG%"=="" set CONFIG=Release

set SCRIPT_DIR=%~dp0
set PROJECT_ROOT=%SCRIPT_DIR%..
set DIST_DIR=%PROJECT_ROOT%\dist
set WINDOWS_POWERSHELL=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe

if not exist "%WINDOWS_POWERSHELL%" (
  echo ERROR: Windows PowerShell is required for VSTO certificate signing.
  exit /b 1
)

echo ============================================
echo  Zotero Linker Installer Build
echo  Version: %VERSION%
echo  Configuration: %CONFIG%
echo ============================================

echo [1/2] Building Word VSTO add-in...
"%WINDOWS_POWERSHELL%" -NoProfile -ExecutionPolicy Bypass -File "%PROJECT_ROOT%\tools\Build-VstoAddIn.ps1" ^
  -Configuration %CONFIG% ^
  -Version %VERSION%
if %ERRORLEVEL% neq 0 (
  echo ERROR: VSTO build failed.
  exit /b 1
)

echo [2/2] Building installer...
if not exist "%DIST_DIR%" mkdir "%DIST_DIR%"

for %%d in ("%ProgramFiles(x86)%\Inno Setup 6" "%ProgramFiles%\Inno Setup 6" "D:\Program Files (x86)\Inno Setup 6") do (
  if exist "%%~d\ISCC.exe" set ISCC=%%~d\ISCC.exe
)
if not defined ISCC (
  for /f "delims=" %%i in ('where iscc 2^>nul') do set ISCC=%%i
)
if not defined ISCC (
  echo ERROR: Inno Setup 6 not found. Install from https://jrsoftware.org/isinfo.php
  exit /b 1
)

"%ISCC%" /DVersion=%VERSION% /DConfig=%CONFIG% "%SCRIPT_DIR%setup.iss"
if %ERRORLEVEL% neq 0 (
  echo ERROR: Installer build failed.
  exit /b 1
)

echo ============================================
echo  Installer built successfully!
echo  Output: %DIST_DIR%\ZoteroLinkerSetup-%VERSION%.exe
echo ============================================
