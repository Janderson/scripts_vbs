on error resume next

Set Shell = CreateObject("WScript.Shell")

Function infonetwork()
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	strin = "" & chr(10)
	For Each objItem in colItems
		strin = strin & objItem.MACAddress & " : "
		For Each strAddress in objItem.IPAddress
			strin = strin & strAddress & chr(10)
		Next
	Next
	infonetwork = strin
end Function
CompName = Shell.ExpandEnvironmentStrings("%COMPUTERNAME%")
UserName = Shell.ExpandEnvironmentStrings("%USERNAME%")
programfolder = Shell.ExpandEnvironmentStrings("%programfiles%")
ipaddr = infonetwork()
Const ForAppending = 2
userprofile = Shell.ExpandEnvironmentStrings("%userprofile%")
Set objFs = CreateObject("Scripting.FileSystemObject")
Set objFile = objFs.OpenTextFile(userprofile & "\abertura.html", ForAppending, True)
objFile.Write "<html><head><style type='text/css'></style>"

objFile.Write "</head><body><h2>Abertura de Chamados</h2>"
objFile.Write "<form action='https://192.3.2.241/ocomon/ocomon/geral/incluir_user_ex.dev.php'>"
objFile.Write "<table border=1>"
objFile.Write "<tr><td>Área Responsável:</td><td><select name='sistemas'><option value='9'>Suporte Sistemas</option><option value='8' selected>Suporte Técnico</option></select></td></tr>"
objFile.Write "<tr><td>Ramal:</td><td><input type='text' name='contato' value='' >"
objFile.Write "<input type='hidden' name='usuario' value='" & UserName & "' >"
objFile.Write "<input type='hidden' name='host' value='" & CompName & "' >"
objFile.Write "<input type='hidden' name='enderecoip' value='" & ipaddr & "' >"
objFile.Write "<input type='hidden' name='salvar' id='salvar' value='True'></td></tr>"
objFile.Write "<tr><td>Descrição do Problema:</td><td><textarea cols=100 rows=5 name='problema'></textarea></td>"
objFile.Write "<tr><td colspan='2'><center><input type='submit' value='Confirmar Chamado'></center></td></table>"
objFile.Write "</form></body></html> "
objFile.Close

Shell.Exec(programfolder & "\Internet Explorer\iexplore.exe " & userprofile & "\abertura.html")