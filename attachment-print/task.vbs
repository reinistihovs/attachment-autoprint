Set WinScriptHost = CreateObject("WScript.Shell")
WinScriptHost.Run Chr(34) & "C:\attachment-print\task.bat" & Chr(34), 0
Set WinScriptHost = Nothing