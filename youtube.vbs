Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c youtube.bat"
oShell.Run strArgs, 0, false