Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("PROCESS")
MsgBox WshSysEnv("USERPROFILE")