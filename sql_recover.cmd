@echo off
title SQL Recover
echo ============================================
echo   PM2014 SQL login recovery (n1 account)
echo ============================================
echo.
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo [ERROR] Not running as administrator.
  echo Right-click the shortcut then Run as administrator.
  echo.
  pause
  exit /b 1
)
echo Running recovery as SYSTEM ...
echo.
schtasks /create /tn PhRec /ru SYSTEM /rl highest /sc once /st 23:59 /f /tr "cmd /c sqlcmd -S .\PM2014 -E -i C:\Users\n1\Desktop\pharm\phrec.sql -o C:\Users\n1\Desktop\pharm\phrec_out.txt" >nul
schtasks /run /tn PhRec >nul
timeout /t 6 /nobreak >nul
schtasks /delete /tn PhRec /f >nul 2>&1
echo ----------- RESULT -----------
if exist C:\Users\n1\Desktop\pharm\phrec_out.txt (
  type C:\Users\n1\Desktop\pharm\phrec_out.txt
) else (
  echo No result file. SYSTEM connection may have failed.
)
echo ------------------------------
echo.
echo   If you see  sysadmin_now = 1  it worked.
echo   Tell Claude what is shown above.
echo.
pause