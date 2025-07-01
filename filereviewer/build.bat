@echo off
REM Build the single-file, windowed EXE
pyinstaller --clean file_report_app.spec
echo.
echo Build complete.  Grab dist\FileReportApp.exe
pause