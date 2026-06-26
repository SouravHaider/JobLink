@echo off
REM ============================================================
REM   JobLink - one-click build for Windows 11 (64-bit)
REM   Put this file in the JobLink project folder and
REM   double-click it. Node.js (LTS) must be installed first:
REM   https://nodejs.org
REM ============================================================
title JobLink - Windows Build

REM Run from this script's own folder (the project root)
cd /d "%~dp0"

echo ============================================
echo   Building JobLink for Windows 11 (64-bit)
echo ============================================
echo.

REM --- Check Node.js is installed ---
where node >nul 2>nul
if %errorlevel% neq 0 (
  echo [ERROR] Node.js was not found.
  echo Install the LTS version from https://nodejs.org then run this again.
  echo.
  pause
  exit /b 1
)

REM --- Check this is the right folder ---
if not exist "package.json" (
  echo [ERROR] package.json not found.
  echo Make sure this file is inside the JobLink project folder.
  echo.
  pause
  exit /b 1
)

echo Step 1 of 2: Installing dependencies (this can take a few minutes)...
echo.
call npm install
if %errorlevel% neq 0 (
  echo.
  echo [ERROR] npm install failed. See the messages above.
  pause
  exit /b 1
)

echo.
echo Step 2 of 2: Building the Windows installer...
echo.
call npm run electron:build:win
if %errorlevel% neq 0 (
  echo.
  echo [ERROR] The build failed. See the messages above.
  pause
  exit /b 1
)

echo.
echo ============================================
echo   Done!  Your installer is here:
echo   release\JobLink Setup 1.0.0.exe
echo ============================================
echo.
echo Opening the release folder...
start "" "release"
pause
