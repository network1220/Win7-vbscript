Set oShell = WScript.CreateObject ("WScript.Shell")


oShell.run "cmd.exe /C shutdown /t 0 /r"

Set oShell = Nothing'