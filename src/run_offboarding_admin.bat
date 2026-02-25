@echo off
setlocal
set "PS1=%~dp0OffboardingAutomation.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
"Start-Process powershell.exe -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -STA -File ""%PS1%""'"
endlocal