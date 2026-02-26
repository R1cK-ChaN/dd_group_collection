@echo off
title DingTalk Group File Collector
echo ========================================
echo  DingTalk Group File Collector
echo  Auto-restart wrapper
echo ========================================

:loop
echo.
echo [%date% %time%] Starting collector...
python "%~dp0run_claude.py"
echo.
echo [%date% %time%] Process exited with code %ERRORLEVEL%.
echo Restarting in 30 seconds... (Ctrl+C to stop)
timeout /t 30 /nobreak
goto loop
