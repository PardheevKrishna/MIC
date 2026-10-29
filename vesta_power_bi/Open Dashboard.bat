@echo off
rem Double-click to open the Vesta dashboard refreshed with the CSVs in data_private\ (or data\).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\open_dashboard.ps1"
if errorlevel 1 pause
timeout /t 5 >nul
