On Error Resume next ' Se voce comentar esse codigo fica em modo debug
' Constantes dos Grupos 
Const BRADESCO_GROUP = "cn=gpo f1 - utilizadores bradesco"
Const BI_GROUP = "cn=gpo f1 - utilizadores bi"
Const GNRE_GROUP = "cn=gpo f1 - utilizadores gnre"
Const PRIMUS_GROUP = "cn=gpo f1 - utilizadores primus"
Const FINANCEIRO_GROUP = "cn=gpo f1 - financeiro"
Const COMPRAS_GROUP = "cn=gpo f1 - compras"
Const CONTABILIDADE_GROUP = "cn=gpo f1 - contabilidade"
Const TI_GROUP = "cn=gpo f1 - ti"
Const FILE_SERVER_NAME = "192.3.2.10"
' Objetos 
Set wshNetwork = CreateObject("WScript.Network") 
Set ADSysInfo = CreateObject("ADSystemInfo") 
Set CurrentUser = GetObject("LDAP://" & ADSysInfo.UserName) 
Set objShell = CreateObject("Shell.Application")
Set oShell = CreateObject("WScript.Shell")

'msgbox CurrentUser.get("username")
' Verifica para qual grupo vai e executa a subrotina
strGroups =LCase(Join(CurrentUser.MemberOf)) ' Coloca todos com letra minuscula para não dar erro

if InStr(strGroups, TI_GROUP) then
	'msgbox "ti"
end if
if InStr(strGroups, COMPRAS_GROUP) then
	compras
end if
if InStr(strGroups, FINANCEIRO_GROUP) then
	financeiro
end if
if InStr(strGroups, CONTABILIDADE_GROUP) then
	contabilidade
end if
if InStr(strGroups, BRADESCO_GROUP) then
	comp_bancos
	icon_bradesco
end if
if InStr(strGroups, GNRE_GROUP) then
	comp_bancos
	icon_gnre
	icon_gfinanceiro
end if
if InStr(strGroups, PRIMUS_GROUP) then
	usa_primus
end if
if InStr(strGroups, BI_GROUP) then
	icon_bi
end if

' Operação para todos os usuarios

todos

' subrotinas de compartilhamentos
sub comp_bancos()
	wshNetwork.MapNetworkDrive "B:", "\\" & FILE_SERVER_NAME & "\bancos"
	objShell.NameSpace("B:").Self.Name = "Bancos"
end sub

sub usa_primus()
	' Cria o compartilhamento do  PRIMUS
	'wshNetwork.RemoveNetworkDrive "R:" ' Disconecta o driver
	wshNetwork.MapNetworkDrive "R:", "\\" & FILE_SERVER_NAME & "\sistemas"
	objShell.NameSpace("R:").Self.Name = "Primus"

	' Cria o Atalho do PRIMUS
	ptdesk = oShell.SpecialFolders("Desktop")
	'msgbox ptdesk & "\PRIMUS.lnk"
	Set objShortCut = oShell.CreateShortcut(ptdesk & "\PRIMUS.lnk")
	objShortCut.TargetPath = "R:\primus_producao\primus\prg\primus.exe"
	objShortCut.Description = "Atalho para o PRIMUS"
	objShortCut.WindowStyle = 3
	objShortCut.HotKey = sHotKey
	objShortCut.IconLocation = "R:\primus_producao\primus\prg\primus.ico"
	objShortCut.WorkingDirectory = "R:\primus_producao\PRIMUS\arq"
	objShortCut.Save

	odbc_name="PRIMUS"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\ODBC Data Sources\" & odbc_name, "PRIMUS ODBC","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\AllowDataCompression","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\AllowProcCalls","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\AllowUnsupportedChar","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\BlockFetch","1","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\BlockSizeKB","32","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\CatalogOptions","3","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\CCSID","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\CommitMode","2","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Concurrency","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\ConnectionType","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\CursorSensitivity","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Database","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\DateFormat","5","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\DateSeparator","1","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Decimal","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\DefaultLibraries","PRIMUS","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\DefaultPackage","PRIMUSPAC/DEFAULT(IBM),2,0,1,0,512","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\DefaultPkgLibrary","PRIMUSPAC","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Description","PRIMUS","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Driver","C:\WINDOWS\system32\cwbodbc.dll","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\ExtendedColInfo","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\ExtendedDynamic","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\ForceTranslation","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Graphic","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\HexParserOpt","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\LanguageID","ENU","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\LazyClose","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\LibraryView","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\MaxFieldLength","32","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\MaximumDecimalPrecision","31","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\MaximumDecimalScale","31","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\MaxTraceSize","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\MinimumDivideScale","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\MultipleTraceFiles","1","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Naming","1","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\ODBCRemarks","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\PreFetch","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\QAQQINILibrary","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\QueryOptimizeGoal","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\QueryTimeout","1","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SearchPattern","1","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Signon","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SortSequence","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SortTable","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SortWeight","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SQDiagCode","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SQLConnectPromptMode","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SSL","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\System","192.1.0.3","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\TimeFormat","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\TimeSeparator","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Trace","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\TraceFileNamet","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\TranslationDLL","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\TranslationOption","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\TrueAutoCommit","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\UnicodeSQL","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\UserID","","REG_SZ"
end sub

sub usa_primus_homo()
	' Cria o compartilhamento do  PRIMUS
	'wshNetwork.RemoveNetworkDrive "R:" ' Disconecta o driver
	wshNetwork.MapNetworkDrive "R:", "\\" & FILE_SERVER_NAME & "\sistemas"
	objShell.NameSpace("R:").Self.Name = "Primus"

	' Cria o Atalho do PRIMUS
	ptdesk = oShell.SpecialFolders("Desktop")
	'msgbox ptdesk & "\PRIMUS.lnk"
	Set objShortCut = oShell.CreateShortcut(ptdesk & "\PRIMUS.lnk")
	objShortCut.TargetPath = "R:\primus_producao\primus\prg\primus.exe"
	objShortCut.Description = "Atalho para o PRIMUS"
	objShortCut.WindowStyle = 3
	objShortCut.HotKey = sHotKey
	objShortCut.IconLocation = "R:\primus_producao\primus\prg\primus.ico"
	objShortCut.WorkingDirectory = "R:\primus_producao\PRIMUS\arq"
	objShortCut.Save
end sub

sub odbc_primus()
	odbc_name="PRIMUS"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\ODBC Data Sources\" & odbc_name, "PRIMUS ODBC","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\AllowDataCompression","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\AllowProcCalls","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\AllowUnsupportedChar","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\BlockFetch","1","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\BlockSizeKB","32","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\CatalogOptions","3","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\CCSID","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\CommitMode","2","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Concurrency","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\ConnectionType","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\CursorSensitivity","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Database","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\DateFormat","5","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\DateSeparator","1","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Decimal","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\DefaultLibraries","PRIMUS","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\DefaultPackage","PRIMUSPAC/DEFAULT(IBM),2,0,1,0,512","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\DefaultPkgLibrary","PRIMUSPAC","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Description","PRIMUS","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Driver","C:\WINDOWS\system32\cwbodbc.dll","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\ExtendedColInfo","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\ExtendedDynamic","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\ForceTranslation","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Graphic","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\HexParserOpt","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\LanguageID","ENU","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\LazyClose","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\LibraryView","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\MaxFieldLength","32","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\MaximumDecimalPrecision","31","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\MaximumDecimalScale","31","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\MaxTraceSize","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\MinimumDivideScale","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\MultipleTraceFiles","1","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Naming","1","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\ODBCRemarks","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\PreFetch","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\QAQQINILibrary","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\QueryOptimizeGoal","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\QueryTimeout","1","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SearchPattern","1","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Signon","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SortSequence","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SortTable","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SortWeight","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SQDiagCode","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SQLConnectPromptMode","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\SSL","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\System","192.1.0.3","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\TimeFormat","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\TimeSeparator","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\Trace","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\TraceFileNamet","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\TranslationDLL","","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\TranslationOption","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\TrueAutoCommit","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\UnicodeSQL","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\UserID","","REG_SZ"
end sub

sub assinatura()
	Set outputfileObject = CreateObject("Scripting.FileSystemObject")
	Set oShell = CreateObject("WScript.Shell")
	Set WshSysEnv = oShell.Environment("PROCESS")
	path = WshSysEnv("USERPROFILE") & "\Dados de aplicativos\Microsoft\Assinaturas"
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
	'msgbox username
	departamento = CurrentUser.Get("Department")
	cargo = CurrentUser.Get("Title")
	fone = CurrentUser.Get("telephoneNumber")
	fax = "3212 - 3455"
	email = CurrentUser.Get("mail")
	
	do while not objtextstream.atendofstream
		rdtxt=objtextstream.readline
		if InStr(rdtxt, "<div class=Section1>") then 
			'msgbox "!!!! encontrado  !!!!!"
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

sub cria_atalho()
	Set objShortCut = objwShell.CreateShortcut("C:\PRIMUS.lnk")
	objShortCut.Description = "PRIMUS"               
	objShortCut.WindowStyle = iWinStyle
	objShortCut.HotKey = sHotKey
	objShortCut.Save
end sub 

sub antivirus() 
	set WshShell = CreateObject("WScript.Shell")
	cmdrun = "runas /user:GUERRA\administrator ""\\192.3.2.10\ofcscan\IpXfer\IpXfer.exe -m 1 -s 192.3.2.10 -p 8080 -c 39438 -q 1"""
	WshShell.Run cmdrun, 0
	WScript.Sleep 2000
	WshShell.Sendkeys "@.guer$a0812~"
	WScript.Sleep 200
	WshShell.Run "runas /user:GUERRA\administrator ""\\192.3.2.10\ofcscan\AutoPcc.exe""", 0
	WScript.Sleep 2000
	WshShell.Sendkeys "@.guer$a0812~"	
end sub

sub atu_dthora()
	oShell.Run "net time \\10.1.1.2 /set /yes", 0, True
end sub 

sub compras() 
	'msgbox "COMPRAS"
end sub
                                                               

' procedimentos executados para todos

sub todos()

	' Verifica se o outlook esta instalado para criar o icone na area de trabalho
	dim filesys
	Set filesys = CreateObject("Scripting.FileSystemObject") 
	if filesys.FileExists("C:\Arquivos de programas\Microsoft Office\Office12\outlook.exe") then
		' Cria icone OUTLOOK 2007 na área de trabalho 
        	Set objwShell = CreateObject("Wscript.Shell")
		ptdesk = objwShell.SpecialFolders("Desktop")
		Set objShortCut = objwShell.CreateShortcut(ptdesk & "\Outlook 2007.lnk")
		objShortCut.TargetPath = "C:\Arquivos de programas\Microsoft Office\Office12\outlook.exe"
		objShortCut.Description = "Atalho para Outlook 2007"
		objShortCut.WindowStyle = iWinStyle
	        objShortCut.Save
	end if

 	' Cria icone na área de trabalho PARTICULAR
        Set objwShell = CreateObject("Wscript.Shell")
	ptdesk = objwShell.SpecialFolders("Desktop")
	Set objShortCut = objwShell.CreateShortcut(ptdesk & "\PARTICULAR.lnk")
	objShortCut.TargetPath = "U:\"
	objShortCut.Description = "Atalho para a unidade particular"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.WorkingDirectory = "U:\"
	objShortCut.Save

	' Atualiza o anti-virus
	antivirus

	' Atualiza data e hora
	atu_dthora

	assinatura

	oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Associations\LowRiskFileTypes","primus.exe,inicia.exe","REG_SZ"

          ' Estabelece o compartilhamento PUBLICA
	wshNetwork.MapNetworkDrive "P:", "\\" & FILE_SERVER_NAME & "\PUBLICA"
	objShell.NameSpace("P:").Self.Name = "Publica"

end sub

sub financeiro() 
	
         ' Estabelece o compartilhamento FINANCEIRO 
	wshNetwork.MapNetworkDrive "S:", "\\" & FILE_SERVER_NAME & "\FINANCEIRO"
	objShell.NameSpace("S:").Self.Name = "Financeiro"
         
         ' Cria icone FINANCEIRO na área de trabalho 
        Set objwShell = CreateObject("Wscript.Shell")
	ptdesk = objwShell.SpecialFolders("Desktop")
	Set objShortCut = objwShell.CreateShortcut(ptdesk & "\FINANCEIRO.lnk")
	objShortCut.TargetPath = "S:\"
	objShortCut.Description = "Atalho para a pasta do setor"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.WorkingDirectory = "S:\"
	objShortCut.Save
         
	 ' Renomeia unidade pessoal 
        objShell.NameSpace("U:").Self.Name = "Particular"

end sub

sub contabilidade() 
	
         ' Estabelece o compartilhamento FINANCEIRO 
	wshNetwork.MapNetworkDrive "S:", "\\" & FILE_SERVER_NAME & "\CONTABILIDADE"
	objShell.NameSpace("S:").Self.Name = "Contabilidade"
         
         ' Cria icone FINANCEIRO na área de trabalho 
        Set objwShell = CreateObject("Wscript.Shell")
	ptdesk = objwShell.SpecialFolders("Desktop")
	Set objShortCut = objwShell.CreateShortcut(ptdesk & "\CONTABILIDADE.lnk")
	objShortCut.TargetPath = "S:\"
	objShortCut.Description = "Atalho para a pasta do setor"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.WorkingDirectory = "S:\"
	objShortCut.Save
         
	 ' Renomeia unidade pessoal 
        objShell.NameSpace("U:").Self.Name = "Particular"

end sub

sub icon_bradesco() 
	
         ' Cria o compartilhamento
	Set objwShell = CreateObject("Wscript.Shell")
	ptdesk = objwShell.SpecialFolders("Desktop")
	Set objShortCut = objwShell.CreateShortcut(ptdesk & "\Office Banking Bradesco.lnk")
	objShortCut.TargetPath = "B:\bradesco\OBB2\INICIA.EXE"
	objShortCut.Arguments = "OBB.EXE"
	objShortCut.Description = "SoftWare Bradesco"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.WorkingDirectory = "B:\bradesco\OBB2\"
	objShortCut.Save

end sub

sub icon_bi()
	
         ' Cria o compartilhamento
	Set objwShell = CreateObject("Wscript.Shell")
	ptdesk = objwShell.SpecialFolders("Desktop")
	Set objShortCut = objwShell.CreateShortcut(ptdesk & "\BI.lnk")
	objShortCut.TargetPath = "http://192.0.0.81/gxplorer50/gxplorer.asp"
	objShortCut.Description = "SoftWare BI"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.Save

end sub

sub icon_gnre() 
         ' Cria o compartilhamento
	Set objwShell = CreateObject("Wscript.Shell")
	ptdesk = objwShell.SpecialFolders("Desktop")
	Set objShortCut = objwShell.CreateShortcut(ptdesk & "\gnre.lnk")
	objShortCut.TargetPath = "B:\brasil\gnre.exe"
	objShortCut.Description = "SoftWare GNRE"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.WorkingDirectory = "b:\brasil"
	objShortCut.Save

end sub

sub icon_gfinanceiro() 
	
         ' Cria o compartilhamento
	Set objwShell = CreateObject("Wscript.Shell")
	ptdesk = objwShell.SpecialFolders("Desktop")
	Set objShortCut = objwShell.CreateShortcut(ptdesk & "\Gerenciador Financeiro.lnk")
	objShortCut.TargetPath = "b:\brasil\bancobrasil\officeIE\index.html"
	objShortCut.IconLocation = "b:\brasil\bancobrasil\officeIE\imagem\bb.ico"
	objShortCut.Description = "Gerenciador Financeiro"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.WorkingDirectory = "b:\brasil\bancobrasil"
	objShortCut.Save

end sub

'oShell.Run "%SystemRoot%\system32\cmd.exe", 1, True
'oShell.Run "runas /usuario:administrador\SAFSBILLING%SystemRoot%\notepad.exe", 1, True
'oShell.Run("runas /user:" & "administrador" & " " & CHR(34) & "admadm" & CHR(34), 2, false)
'objShell.ShellExecute "notepad.exe", "administrador" , "admadm", "runas", 1
' Compartilhamento do bradesco
