set ADSysInfo = CreateObject("ADSystemInfo") 
set objUser = GetObject("LDAP://" & ADSysInfo.UserName)
set oShell = CreateObject("WScript.Shell")
' Busca os grupos e joga na variavel arrMemberof
arrMemberOf = objUser.GetEx("memberOf")

For Each Group in arrMemberOf
   ' Separa so o nome do grupo e executa os scripts na pasta grupos
   a=Split(Group,",")
   b=Split(a(0),"=")
   oShell.run "cscript ""\\192.3.2.62\netlogon\grupos\" & b(1) & ".vbs""", 0
Next