Set ADSysInfo = CreateObject("ADSystemInfo") 
Set CurrentUser = GetObject("LDAP://" & ADSysInfo.UserName) 

sub assinatura()
'Create File Object
	Set outputfileObject = CreateObject("Scripting.FileSystemObject")
	Set oShell = CreateObject("WScript.Shell")
	Set WshSysEnv = oShell.Environment("PROCESS")
	path = WshSysEnv("USERPROFILE") & "\Dados de aplicativos\Microsoft\Assinaturas"
	msgbox path
	if not outputfileObject.FolderExists(path) then 
		outputfileObject.CreateFolder(path)
		outputfileObject.CreateFolder(path & "\padrao_arquivos")
		outputfileObject.CopyFile "\\192.3.2.10\publica\Assinaturas\padrao_arquivos\*", path & "\padrao_arquivos\", True
	end if
	
	Set WshSysEnv =oShell.Environment("PROCESS")
	
	Set objfileobject = CreateObject("Scripting.FilesystemObject")
	Set objFile = objfileobject.CreateTextFile(path & "\padrao.htm")
	Set objfiletoread = objfileobject.getfile("\\192.3.2.10\publica\Assinaturas\padrao.htm")
	Set objtextstream = objfiletoread.openastextstream
	username = CurrentUser.Get("displayName") 
	departamento = "Departamento de Engª Industrial"
	cargo = "Supervisor Eng.ª Processo"
	fone = CurrentUser.Get("telephoneNumber")
	fax = "32123455"
	email = CurrentUser.Get("mail")
	
	do while not objtextstream.atendofstream
		rdtxt=objtextstream.readline
		if InStr(rdtxt, "<div class=Section1>") then ' pode ser strcomp para ficar mais preciso
			msgbox "!!!! encontrado  !!!!!"
			objFile.writeline("<div class=Section1>")
			objFile.writeline("<p class=MsoAutoSig>" & uCase(username) & "</p>")
			objFile.writeline("<p class=MsoAutoSig><o:p>&nbsp;</o:p></p>")
			objFile.writeline("<p class=MsoAutoSig>" & "Guerra S.A. Implementos Rodoviários" & "</p>")
			objFile.writeline("<p class=MsoAutoSig>" & departamento & "</p>")
			objFile.writeline("<p class=MsoAutoSig>" & cargo & "</p>")
			objFile.writeline("<p class=MsoAutoSig>End: Br 116 km 146,4     Bairro Mariland</p>")
			objFile.writeline("<p class=MsoAutoSig>CEP: 95059-520   Caxias do Sul - RS - Brasil</p>")
			objFile.writeline("<p class=MsoAutoSig>Fone:" & fone & "</p>")
			objFile.writeline("<p class=MsoAutoSig>Fax:" & fax & "</p>")
			objFile.writeline("<p class=MsoAutoSig>email:" & email & "</p>")
		else 
			objFile.writeline(rdtxt)
		end if
	loop
	objFile.close

end sub
assinatura