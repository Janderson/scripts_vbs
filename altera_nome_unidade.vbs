' Para qualquer erro continue
On Error Resume Next
' OBJETOS GLOBAIS

Set obShell = CreateObject("WScript.Shell")
Set objShell = CreateObject("Shell.Application")
Set objNetwork = CreateObject("Wscript.Network")
strUserName = objNetwork.UserName

objShell.NameSpace("S:").Self.Name = "Financeiro"
objShell.NameSpace("P:").Self.Name = "Publica"
objShell.NameSpace("U:").Self.Name = strUserName
