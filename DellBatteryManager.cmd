@echo off
rem Lance Dell Battery Manager. Le script demande lui-meme les droits administrateur (UAC).
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0DellBatteryManager.ps1"
