@echo off
setlocal
pushd %~dp0
set VENV=.venv
rem stop running exe to avoid dist cleanup lock
taskkill /IM naos_sla_ui.exe /F /T >nul 2>&1
taskkill /IM naos_sla.exe /F /T >nul 2>&1
timeout /t 1 >nul 2>&1
if not exist %VENV% (
    python -m venv %VENV%
)
call %VENV%\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller
rem clean dist to avoid permission issues; ignore failures
if exist dist\naos_sla (
    rmdir /s /q dist\naos_sla >nul 2>&1
)
pyinstaller --clean --noconfirm naos_sla.spec
echo Build finished. Distribution in .\dist\naos_sla
popd
