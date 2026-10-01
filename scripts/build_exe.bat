@echo off
rem Build a standalone pst_testgen.exe with PyInstaller (dist\pst_testgen.exe).
rem The executable still needs classic desktop Outlook on the machine it runs on.
rem
rem Needs Python 3.10+ (py launcher or python on PATH).
rem Optional: set DIST / WORK to change the output and scratch folders.

setlocal
cd /d "%~dp0.."
if "%DIST%"=="" set "DIST=dist"
if "%WORK%"=="" set "WORK=build"

rem PY can be set by the caller (CI sets PY=python)
if "%PY%"=="" set "PY=py -3"
%PY% --version >nul 2>&1 || set "PY=python"
%PY% --version >nul 2>&1 || (echo ERROR: Python 3.10+ not found. & exit /b 1)

echo [1/3] Creating build environment in %WORK%\venv ...
if exist "%WORK%\venv" rmdir /s /q "%WORK%\venv"
%PY% -m venv "%WORK%\venv" || exit /b 1
set "VPY=%WORK%\venv\Scripts\python.exe"
"%VPY%" -m pip install --quiet --upgrade pip
"%VPY%" -m pip install --quiet ".[build]" || (echo ERROR: pip install failed. & exit /b 1)

echo [2/3] Building pst_testgen.exe ...
"%VPY%" -m PyInstaller --noconfirm --clean --onefile --console ^
    --name pst_testgen --distpath "%DIST%" --workpath "%WORK%\pst_testgen" --specpath "%WORK%" ^
    --hidden-import win32com.mapi.mapi --hidden-import pythoncom --hidden-import pywintypes ^
    --hidden-import win32timezone ^
    "packaging\entry_pst_testgen.py" || (echo ERROR: build failed. & exit /b 1)

echo [3/3] Smoke test ...
"%DIST%\pst_testgen.exe" --version || exit /b 1

echo.
echo Done: %DIST%\pst_testgen.exe
endlocal
