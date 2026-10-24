Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\Lenovo\Desktop\updates-app"
WshShell.Run "cmd /c start-app.bat", 0, False
