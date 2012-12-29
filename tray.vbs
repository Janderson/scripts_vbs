Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
cbSize As Long
hWnd As Long
uID As Long
uFlags As Long
uCallbackMessage As Long
hIcon As Long
szTip As String * 64

End Type

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Public Enum Acoes
NIM_ADD = &H0
NIM_DELETE = &H2
NIM_MODIFY = &H1

End Enum

Public Sub Tray(ByVal hWnd&, ByVal hIcon As StdPicture, ByVal sTip$, Acao As Acoes)

Dim i As NOTIFYICONDATA

With i
.hWnd = hWnd
.hIcon = hIcon
.szTip = sTip & vbNullChar
.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
.uCallbackMessage = &H200
.uID = vbNull
.cbSize = Len(i)

Shell_NotifyIcon Acao, i

End With

End Sub

' TRAY

'Para adicionar 1 ícone no tray faca o seguinte
Tray me.hwnd, app.icon,"Mensagem do tooltip",NIM_ADD

'Para remover o icone do tray faca o seguinte
Tray me.hwnd, app.icon, "Mensagem do tooltip", NIM_DELETE

'Para modificar quaisquer informacoes sobre do tray faca o seguinte
Tray me.hwnd, app.icon, "Nova mensagem do tooltip", NIM_MODIFY

Onde:
Me.hwnd = handle do form
App.Icon = ícone q será mostrado no tray
"Mensagem do tooltip" = mensagem que aparece qd c posiciona o mouse por alguns segundos sobre o icone
NIM_MODIFY, NIM_ADD, NIM_DELETE = acoes de adicionar/modificar/remover icone do tray