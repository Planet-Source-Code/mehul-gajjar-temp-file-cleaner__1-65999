VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "TmpCleaner"
   ClientHeight    =   5175
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   4485
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":2CFA
   ScaleHeight     =   5175
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   3
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   4755
      TabIndex        =   6
      Top             =   2690
      Width           =   4755
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "This code delete all tmp file  from Temp Folder  (not including sub folder)  when  you shutdown your pc......"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   4335
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4755
      Width           =   4095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Load at startup"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Tag             =   "LST"
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   4200
      Top             =   6240
   End
   Begin VB.FileListBox File2 
      Height          =   285
      Left            =   8760
      Pattern         =   "*.exe"
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   8760
      Pattern         =   "*.tmp"
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   3405
      Left            =   -240
      TabIndex        =   4
      Top             =   -480
      Width           =   5175
      ExtentX         =   9128
      ExtentY         =   6006
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      X1              =   240
      X2              =   4320
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      X1              =   240
      X2              =   4320
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   240
      X2              =   4320
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   240
      X2              =   4320
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Minimize this program and use with system tray."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   3600
      Width           =   3495
   End
   Begin VB.Menu ppme 
      Caption         =   "ppmenu"
      Visible         =   0   'False
      Begin VB.Menu sshow 
         Caption         =   "Show"
      End
      Begin VB.Menu eend 
         Caption         =   "End"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Dim tfol As String
Dim strText As String
Dim showfrm As Boolean



Private Sub Check1_Click()
If File2.List(0) = "" Then
Check1.Value = 0
MsgBox "Please Make TmpCleaner.exe , After user Load at startup", vbInformation, "MSoft"
Exit Sub
Else
      If Check1.Value = 1 Then
         wscr.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\TmpCleaner", App.Path & "\" & File2.List(0), "REG_SZ"
         wscr.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS" & "\users\LST", 1, "REG_DWORD"
      Else
         On Error Resume Next
         wscr.regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\TmpCleaner"
         wscr.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS" & "\users\LST", 0, "REG_DWORD"
      End If
End If
End Sub

Private Sub eend_Click()
Form_Unload (Cancel)
End Sub

Private Sub Form_Load()
Dim buffer As String, Length As Integer
buffer = Space$(512)
Length = GetTempPath(Len(buffer), buffer)
tfol = Left$(buffer, Length)
File1.Path = tfol
File2.Path = App.Path

Me.Width = 4605

Set wscr = CreateObject("Wscript.SHELL")
Check1.Value = wscr.regread("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS" & "\users\" & Check1.Tag)
ShowSisTrayIcon Me
'Me.Hide

strText = String(30, " ") + "Developed by - Mehul Gajjar , Contect : 35,Akshar duplex, kalol (N.G) Pin - 382 721, Mail : mehulrgajjar@yahoo.co.in "
       
 WebBrowser1.Navigate App.Path & "\Msoftlg.gif"
   
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmp As Long
If Me.WindowState = vbMinimized Or Me.Visible = False Then
    tmp = X \ Screen.TwipsPerPixelX
    If tmp = &H201 Or tmp = &H204 Then
PopupMenu ppme, , , , sshow
If eend.Checked Then
        HideSisTrayIcon
           End
        ElseIf sshow.Checked Then
            With Me
                .WindowState = vbNormal
                .Visible = True
                .Refresh
               
            End With
End If
End If
End If
End Sub

Private Sub Form_Resize()
If showfrm = True Then
Me.WindowState = 0
showfrm = False
Else
Me.Visible = (Me.WindowState <> vbMinimized)
End If
End Sub

Private Sub Form_Terminate()
On Error Resume Next

Dim buffer As String, Length As Integer
On Error Resume Next
buffer = Space$(512)
Length = GetTempPath(Len(buffer), buffer)
tfol = Left$(buffer, Length)
File1.Path = tfol
File1.Refresh

While File1.ListIndex < File1.ListCount - 1
File1.ListIndex = File1.ListIndex + 1
Kill (tfol & File1)
Wend

HideSisTrayIcon
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Dim buffer As String, Length As Integer
On Error Resume Next
buffer = Space$(512)
Length = GetTempPath(Len(buffer), buffer)
tfol = Left$(buffer, Length)
File1.Path = tfol
File1.Refresh


While File1.ListIndex < File1.ListCount - 1
File1.ListIndex = File1.ListIndex + 1
tf = tfol & File1
Kill tf
Wend

HideSisTrayIcon
End
End Sub

Private Sub sshow_Click()
showfrm = True
Me.Show
End Sub

Private Sub Timer1_Timer()
strText = Mid(strText, 2) & Left(strText, 1)
Text1 = strText
End Sub
