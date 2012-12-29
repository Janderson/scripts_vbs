On Error Resume next
Const BRADESCO_GROUP = "cn=gpo f1 - utilizadores bradesco"
Set wshNetwork = CreateObject("WScript.Network") 
Set ADSysInfo = CreateObject("ADSystemInfo") 
Set CurrentUser = GetObject("LDAP://" & ADSysInfo.UserName) 
Set objShell = CreateObject("Shell.Application")
Dim strGroups 
strGroups = split(CurrentUser.MemberOf, ",")
' GRUPO DO BRADESCO
Dim i, c
i = split("83 104 105 118 97 110 115"," ")
For c=lBound(strGroups) To uBound(strGroups)
    msgbox strGroups(c)
Next