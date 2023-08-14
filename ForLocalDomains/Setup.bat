@echo off
pushd "%~dp0"
PowerShell.exe -ExecutionPolicy Bypass -File ".\FromLocalDomain.ps1"
popd
