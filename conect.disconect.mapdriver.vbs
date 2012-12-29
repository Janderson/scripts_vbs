Option Explicit
'-----------------------------------------------------------------------------------------
' NAME: LogonScript
'-----------------------------------------------------------------------------------------
' Step1-Map network drives for a user
'-----------------------------------------------------------------------------------------
Sub MapNetworkDrives
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
Dim APP1, APP3
Const HKCU = &H80000001 'Constant that represent the HKCU Hive

APP1 = "\\server1\"
APP3 = "\\server2\"

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

'Disconnect Production mapped drives
Set objNetwork = WScript.CreateObject("WScript.Network")
objNetwork.RemoveNetworkDrive "G:"
objNetwork.RemoveNetworkDrive "H:"
objNetwork.RemoveNetworkDrive "I:"
objNetwork.RemoveNetworkDrive "J:"
objNetwork.RemoveNetworkDrive "L:"
objNetwork.RemoveNetworkDrive "S:"
objNetwork.RemoveNetworkDrive "X:"

' 'If there's a remembered location (persistent mapping) delete the associated HKCU registry key
' If bolFoundExisting <> True Then
'   Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
'   objReg.GetStringValue HKCU, "Network\" & Left(strLocalDrive, 1), "RemotePath", strShareConnected
'   If strShareConnected <> "" Then
'     objReg.DeleteKey HKCU, "Network\" & Left(strLocalDrive, 1)
'     Set objReg = Nothing
'     bolFoundRemembered = True
'   End If
' End If

'Give the PC time to do the disconnect, wait 300 milliseconds
wscript.sleep 300

'Now check for group memberships and map appropriate drives
For Each objGroup In objUser.Groups
    'Get the group name and change it to lowercase
    'Note: This can be either Lowercase (LCase) or Uppercase (UCase)
    strLowerCaseGroupName = LCase (objGroup.Name) 
    Select Case strLowerCaseGroupName
    'Check for group memberships and take needed action
    
        Case "it-central-admin" 
        
              objNetwork.MapNetworkDrive "I:", APP1 & strUserName & "$", True   
                 objNetwork.MapNetworkDrive "J:", APP3 & "CADMINAPPS$", True
              objNetwork.MapNetworkDrive "S:", APP1 & "CADMIN$", True
              objNetwork.MapNetworkDrive "X:", APP3 & "APPS", True
              
        Case "it-network-support" 
        
              objNetwork.MapNetworkDrive "I:", APP1 & strUserName & "$", True
                 objNetwork.MapNetworkDrive "S:", APP1 & "groups", True
              objNetwork.MapNetworkDrive "X:", APP1 & "apps", True
              
       End Select
    Next 

End Sub
'END CALL To MAP DRIVES

Call MapNetworkDrives
 
