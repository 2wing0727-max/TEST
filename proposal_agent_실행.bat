@echo off
chcp 65001 >nul
cd /d "%~dp0"
python proposal_agent.py %*
echo.
pause
