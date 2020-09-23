Attribute VB_Name = "Tray"
Const WM_MOUSEMOVE = &H200
Const NIF_ICON = &H2
Const NIF_MESSAGE = &H1
Const NIF_TIP = &H4
Const NIM_ADD = &H0
Const NIM_DELETE = &H2
Const MAX_TOOLTIP As Integer = 64

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TOOLTIP
End Type
Dim nfIconData As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public wscr
Public regKeys(45) As String


Sub ShowSisTrayIcon(frm As Form)
With nfIconData
    .hwnd = frm.hwnd
    .uID = frm.Icon
    .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = frm.Icon.Handle
    .szTip = "Tmp file Monitoring" & vbNullChar
    .cbSize = Len(nfIconData)
End With
Call Shell_NotifyIcon(NIM_ADD, nfIconData)
End Sub

Sub HideSisTrayIcon()
Shell_NotifyIcon NIM_DELETE, nfIconData
End Sub

