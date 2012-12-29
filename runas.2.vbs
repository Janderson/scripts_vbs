Dim oShell 
set oShell= Wscript.CreateObject("WScript.Shell") 
Set objSh = CreateObject("Shell.Application")
objSh.ShellExecute "notepad.exe", "c:\autoexec.bat" , "", "runas", 1