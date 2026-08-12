Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "D:\WorkLog"
WshShell.Run ".\.venv\Scripts\python.exe -m tools.discord_bot", 0, False
