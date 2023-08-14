@echo off
pushd "%~dp0"
PowerShell.exe "ipconfig /flushdns" | PowerShell.exe -ExecutionPolicy Bypass -File ".\FromExtDomain.ps1"
popd
