Option Explicit
'-----------------------------------------------------------------------------------------
' NAME: LogonScript
'-----------------------------------------------------------------------------------------
' Step1-Map network drives for a user
'-----------------------------------------------------------------------------------------
Sub groups
	ON ERROR RESUME Next
	Dim objShell 'This is the shell object
	Dim objNetwork ' This is our network object 
	Dim objDomain 'This is our  Domain Object
	Dim strDomainName 'This is our Domain Name 
	Dim strUserName 'This is our user string
	Dim objUser 'This is our UserObject
	Dim objGroup 'This is our group object
	Dim strLowerCaseGroupName 'This String holds the Lower Case group Name
	Dim clDrives 'Collection of Drives
	Dim strShareConnected '
	Dim strLocalDrive 'Name of Local Drive
	Dim bolFoundExisting ' 
	Dim bolFoundRemembered 'Bolean variable thru which drive is remembered
	Dim i 'Array Counter
	Dim objReg 'Registry for mapped drive
	'Create an instance of the Shell
	Set objShell = CreateObject("WScript.Shell")
	'Create an instance of the Network
	Set objNetwork = CreateObject("WScript.Network")
	'Automatically find the domain name
	Set objDomain = GetObject("LDAP://rootDse")
	strDomainName = objDomain.Get("dnsHostName")
	'Grab the user name
	strUserName = objNetwork.UserName
	'Bind to the user object to get user name and check for group memberships later
	Set objUser = GetObject("WinNT://" & strDomainName & "/" & strUserName)
	Dim strDrive 'Temp variable for storing drive letters
	Set objNetwork = WScript.CreateObject("WScript.Network")
	'Now check for group memberships and map appropriate drives
	For Each objGroup In objUser.Groups
	    'msgbox LCase(objGroup.Name) 
	Next
	set objNetwork = WScript.CreateObject("WScript.Network")
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	if (ObjFSO.DriveExists("\\192.3.2.10\ct")) then
		msgbox "existe"
	end if
  
End Sub
'END CALL To MAP DRIVES
Set objUser = GetObject _
  ("LDAP://cn=GPO F1 - ,ou=informatica,dc=GUERRA")
objUser.GetInfo

strProfilePath = objUser.Get("profilePath")
msgbox strProfilePath 
Call groups