Set objADSysInfo = CreateObject("ADSystemInfo") 
strUserDN = objADSysInfo.username 
msgbox "DN:................ " & strUserDN


' Bind to user object 
Set objUser = Getobject("LDAP://" & strUserDN) 
'msgbox "UserID:............ " & objUser.Get("name") 
'msgbox "First name:........ " & objUser.Get("givenName") 
'msgbox "Initials:.......... " & objUser.Get("initials") 
msgbox "Last name:......... " & objUser.Get("sn") 
msgbox "Display name:...... " & objUser.Get("displayName") 
'msgbox "Description:....... " & objUser.Get("description") 
'msgbox "Office:............ " & objUser.Get("physicalDeliveryOfficeName")
msgbox "Telephone number:.. " & objUser.Get("telephoneNumber")
msgbox "E-Mail:............ " & objUser.Get("mail")
msgbox "Endereco:............ " & CurrentUser.Get("streetAddress")
msgbox "CEP:............ " & CurrentUser.Get("postalCode")
