@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1/2] ตรวจสอบ Source...
go test ./...
if errorlevel 1 (
  echo TEST FAILED
  pause
  exit /b 1
)
echo [2/2] เปิด Development Server...
go run .
pause
