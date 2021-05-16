@echo off
cd %~dp0
powershell.exe -ExecutionPolicy Unrestricted "..\src\list_outlook_folders.ps1"
pause
