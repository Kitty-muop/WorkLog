@echo off
echo Installing WorkLog Discord Bot Auto-Start...
copy /Y "D:\WorkLog\tools\run_bot.vbs" "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\WorkLogBot.vbs"
echo [OK] WorkLog Bot will now automatically start hidden in background when Windows boots up!
pause
