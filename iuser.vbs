Set ADSysInfo = CreateObject("ADSystemInfo") 
Set CurrentUser = GetObject("LDAP://" & ADSysInfo.UserName) 
msgbox "Office:............ " & CurrentUser.Get("physicalDeliveryOfficeName")
