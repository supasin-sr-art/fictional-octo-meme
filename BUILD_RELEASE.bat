@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1/2] Running tests...
go test ./...
if errorlevel 1 (
  echo BUILD CANCELLED: TEST FAILED
  pause
  exit /b 1
)
echo [2/2] Building Windows executable...
go build -trimpath -ldflags="-s -w" -o BLACKWOLF_Web_Server.exe .
if errorlevel 1 (
  echo BUILD FAILED
  pause
  exit /b 1
)
echo BUILD SUCCESS: BLACKWOLF_Web_Server.exe
pause
