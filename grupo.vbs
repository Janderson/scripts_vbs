Const BRADESCO_GROUP = "cn=gpo f1 - utilizadores bradesco"
Set wshNetwork = CreateObject("WScript.Network") 
Set ADSysInfo = CreateObject("ADSystemInfo") 
Set CurrentUser = GetObject("LDAP://" & ADSysInfo.UserName) 
Dim objGroup
Dim objDomain
Dim objNetwork
Dim strDomainName 'This is our Domain Name 
Dim strUserName 'This is our user string

Set objDomain = GetObject("LDAP://rootDse")
strDomainName = objDomain.Get("dnsHostName")
msgbox objDomain.Get("dnsHostName")
Set objShell = CreateObject("WScript.Shell")

'Create an instance of the Network
Set objNetwork = CreateObject("WScript.Network")

'Automatically find the domain name
Set objDomain = GetObject("LDAP://rootDse")
strDomainName = objDomain.Get("dnsHostName")

'Grab the user name
strUserName = objNetwork.UserName
'Grab the user name
strUserName = objNetwork.UserName

'Bind to the user object to get user name and check for group memberships later
Set objUser = GetObject("WinNT://" & strDomainName & "/" & strUserName)

strGroups = LCase(Join(CurrentUser.MemberOf))
' GRUPO DO BRADESCO
For Each objGroup In objUser.Groups
	msgbox objGroup
next
'msgbox strGroups
If InStr(strGroups, BRADESCO_GROUP) Then
  wshNetwork.MapNetworkDrive "b:", "\\gcxfs01\bradesco\" 
  objShell.NameSpace("b:").Self.Name = "Bradesco"
end if