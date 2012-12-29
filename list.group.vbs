Const E_ADS_PROPERTY_NOT_FOUND = &h8000500D
Set ADSysInfo = CreateObject("ADSystemInfo") 
Set objOU = GetObject("LDAP://" & ADSysInfo.UserName) 
ObjOU.Filter= Array("user")
For Each objUser in objOU
    msgbox objUser.cn & " is a member of: " 
    msgbox vbTab & "Primary Group ID: " & objUser.Get("primaryGroupID")
    arrMemberOf = objUser.GetEx("memberOf")
    If Err.Number <>  E_ADS_PROPERTY_NOT_FOUND Then
        For Each Group in arrMemberOf
        msgbox vbTab & Group
        Next
    Else
        msgbox vbTab & "memberOf attribute is not set"
        Err.Clear
      End If
    msgbox VbCrLf
Next