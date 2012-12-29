Set ADSysInfo = CreateObject("ADSystemInfo") 
'msgbox ADSysInfo.UserName
Set objUser = GetObject("LDAP://" & ADSysInfo.UserName)
'msgbox objUser.Get("homeDirectory")
'objUser.Put "homeDirectory", "//gcxfs01/publica"
'objUser.SetInfo
msgbox "Bom dia sr. " &objUser.Get("profilePath") & " Sua conta foi desabilitada por motivos de força maior"