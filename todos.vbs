On Error Resume next ' Se voce comentar esse codigo fica em modo debug
' Constantes dos Grupos 

Const GNRE_GROUP = "cn=gpo f1 - utilizadores gnre"
Const BRADESCO_GROUP = "cn=gpo f1 - utilizadores bradesco"
Const BI_GROUP = "cn=gpo f1 - utilizadores bi"
Const PRIMUS_GROUP = "cn=gpo f1 - utilizadores primus"
' Grupos 
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

' Verifica para qual grupo vai e executa a subrotina
strGroups =LCase(Join(CurrentUser.MemberOf)) ' Coloco todos com letra minuscula para não dar erro

if InStr(strGroups, TI_GROUP) then
	'msgbox "ti"
end if
if InStr(strGroups, COMPRAS_GROUP) then
	compras
end if
if InStr(strGroups, FINANCEIRO_GROUP) then
	financeiro
end if
if InStr(strGroups, GNRE_GROUP) then
	comp_bancos
	icone_gnre
end if

if InStr(strGroups, BI_GROUP) then
	comp_bancos
	icone_BI
end if

if InStr(strGroups, BRADESCO_GROUP) then
	comp_bancos
	icone_bradesco
end if
if InStr(strGroups, PRIMUS_GROUP) then
	usa_primus
end if

' Operação para todos os usuarios

  antivirus
  atu_dthora
  user_block



' subrotinas de compartilhamentos
sub comp_bancos()
	wshNetwork.MapNetworkDrive "B:", "\\" & FILE_SERVER_NAME & "\BANCOS"
	objShell.NameSpace("B:").Self.Name = "Bancos"
end sub

sub user_block()
  oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Group Policy Objects\{052E4B99-F9BE-44F3-B295-8240DE4A3C2F}User\Software\Microsoft\Windows\CurrentVersion\Policies\Associations\LowRiskFileTypes", "INICIAL.exe,PRIMUS.EXE", "REG_SZ"

end sub

sub icone_gnre()
	' Cria icone gerenciador financeiro
	ptdesk = oShell.SpecialFolders("Desktop")
	'msgbox ptdesk & "\PRIMUS.lnk"
	Set objShortCut = oShell.CreateShortcut(ptdesk & "\Gerenciador_Financeiro.lnk")
	objShortCut.TargetPath = "C:\BancoBrasil\officeIE\index.html"
	objShortCut.Description = "Gerenciador Financeiro"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.HotKey = sHotKey
	objShortCut.IconLocation = "C:\BancoBrasil\officeIE\imagen\bb.ico"
	objShortCut.WorkingDirectory = "C:\BancoBrasil\officeIE\"
	objShortCut.Save

	Set objShortCut = oShell.CreateShortcut(ptdesk & "\GNRE.lnk")
	objShortCut.TargetPath = "B:\brasil\GNRE.exe"
	objShortCut.Description = "Gerenciador Financeiro"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.HotKey = sHotKey
	objShortCut.IconLocation = "B:\brasil\GNRE.exe"
	objShortCut.WorkingDirectory = "B:\brasil\"
	objShortCut.Save
end sub

sub icone_bradesco()
	' Cria icone gerenciador financeiro
	ptdesk = oShell.SpecialFolders("Desktop")
	'msgbox ptdesk & "\PRIMUS.lnk"
	Set objShortCut = oShell.CreateShortcut(ptdesk & "\Office Banking Bradesco.lnk")
	objShortCut.TargetPath = "B:\bradesco\OBB2\INICIA.EXE"
	objShortCut.Arguments = "OBB.EXE"
	objShortCut.Description = "Office Banking Bradesco"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.HotKey = sHotKey
	objShortCut.IconLocation = "B:\bradesco\OBB2\OBB.EXE"
	objShortCut.WorkingDirectory = "B:\bradesco\OBB2\"
	objShortCut.Save
end sub

sub icone_hsbc()
	' Cria icone gerenciador financeiro
	ptdesk = oShell.SpecialFolders("Desktop")
	Set objShortCut = oShell.CreateShortcut(ptdesk & "\Gerenciador Financeiro.lnk")
	objShortCut.TargetPath = "C:\BancoBrasil\officeIE\index.html"
	objShortCut.Description = "Gerenciador Financeiro"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.IconLocation = "C:\BancoBrasil\officeIE\imagen\bb.ico"
	objShortCut.WorkingDirectory = "C:\BancoBrasil\officeIE\"
	objShortCut.Save
end sub

sub icone_BI()
	' Cria o registro
	oShell.RegWrite "HKCU\Software\microsoft\windows\currentVersion\Internet Settings\Zonemap\ranges\range1\*","2","REG_DWORD"
	oShell.RegWrite "HKCU\Software\microsoft\windows\currentVersion\Internet Settings\Zonemap\ranges\range1\:Range","http:\\192.0.0.81","REG_SZ"
	' Cria icone gerenciador financeiro
	ptdesk = oShell.SpecialFolders("Desktop")
	Set objShortCut = oShell.CreateShortcut(ptdesk & "\BI.lnk")
	objShortCut.TargetPath = "http://192.0.0.81/gxplorer50/gxplorer.asp"
	objShortCut.Description = "BI"
	objShortCut.WindowStyle = iWinStyle
'	objShortCut.IconLocation = "C:\BancoBrasil\officeIE\imagen\bb.ico"
'	objShortCut.WorkingDirectory = "C:\BancoBrasil\officeIE\"
	objShortCut.Save
end sub

sub usa_primus()
	' Cria o compartilhamento do  PRIMUS
	'wshNetwork.RemoveNetworkDrive "R:" ' Disconecta o driver
	wshNetwork.MapNetworkDrive "R:", "\\" & FILE_SERVER_NAME & "\sistemas"
	objShell.NameSpace("R:").Self.Name = "Primus"

	' Cria o Atalho do PRIMUS
	ptdesk = oShell.SpecialFolders("Desktop")
	Set objShortCut = oShell.CreateShortcut(ptdesk & "\PRIMUS.lnk")
	objShortCut.TargetPath = "R:\primus_producao\primus\prg\primus.exe"
	objShortCut.Description = "Atalho para o PRIMUS"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.HotKey = sHotKey
	objShortCut.IconLocation = "R:\primus_producao\primus\prg\primus.ico"
	objShortCut.WorkingDirectory = "R:\primus_producao\PRIMUS\arq"
	objShortCut.Save

	odbc_name="PRIMUS"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\ODBC Data Sources\" & odbc_name,"PRIMUS ODBC","REG_SZ"
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
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\ TrueAutoCommit","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\UnicodeSQL","0","REG_SZ"
	oShell.RegWrite "HKCU\Software\ODBC\ODBC.INI\"& odbc_name &"\UserID","","REG_SZ"
end sub

sub cria_atalho()
	Set objShortCut = objwShell.CreateShortcut("C:\PRIMUS.lnk")
	objShortCut.Description = "PRIMUS"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.HotKey = sHotKey
	objShortCut.Save
end sub 

sub antivirus() 
	oShell.Run "\\192.3.2.10\publica\IpXfer\IpXfer.exe -m 1 -s 192.3.2.10 -p 8080 -c 39438 -q 1", 0, True
	oShell.Run "\\192.3.2.10\ofcscan\AutoPcc.exe", 0, True
end sub

sub atu_dthora()
	oShell.Run "net time \\10.1.1.2 /set /yes", 0, True
end sub 

sub compras() 
	'msgbox "COMPRAS"
end sub

sub financeiro() 
	' Cria o compartilhamento
	wshNetwork.MapNetworkDrive "S:", "\\" & FILE_SERVER_NAME & "\FINANCEIRO"
	objShell.NameSpace("S:").Self.Name = "Financeiro"

	Set objwShell = CreateObject("Wscript.Shell")
	ptdesk = objwShell.SpecialFolders("Desktop")
	Set objShortCut = objwShell.CreateShortcut(ptdesk & "\FINANCEIRO.lnk")
	objShortCut.TargetPath = "s:\"
	objShortCut.Description = "Atalho para a pasta do setor"
	objShortCut.WindowStyle = iWinStyle
	objShortCut.WorkingDirectory = "S:\"
	objShortCut.Save

end sub

'oShell.Run "%SystemRoot%\system32\cmd.exe", 1, True
'oShell.Run "runas /usuario:administrador\SAFSBILLING%SystemRoot%\notepad.exe", 1, True
'oShell.Run("runas /user:" & "administrador" & " " & CHR(34) & "admadm" & CHR(34), 2, false)
'objShell.ShellExecute "notepad.exe", "administrador" , "admadm", "runas", 1
' Compartilhamento do bradesco
