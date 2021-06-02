@echo off
cd %~dp0
powershell.exe -ExecutionPolicy Bypass "..\src\list_outlook_folders.ps1"
pause
