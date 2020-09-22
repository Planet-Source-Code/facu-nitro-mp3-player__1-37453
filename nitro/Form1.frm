VERSION 5.00
Begin VB.Form tinywindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleWidth      =   1440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   7
      Left            =   960
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   0
      Width           =   240
   End
   Begin VB.Timer nexo 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   -240
      Top             =   -240
   End
   Begin VB.TextBox fileE 
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox listy 
      ForeColor       =   &H00CC9900&
      Height          =   1620
      ItemData        =   "Form1.frx":014A
      Left            =   0
      List            =   "Form1.frx":014C
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.Timer up 
      Interval        =   10
      Left            =   -120
      Top             =   -240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   6
      Left            =   1200
      Picture         =   "Form1.frx":014E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   5
      Left            =   720
      Picture         =   "Form1.frx":0298
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   2100
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   4
      Left            =   480
      Picture         =   "Form1.frx":03E2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   2100
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   720
      Picture         =   "Form1.frx":052C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   480
      Picture         =   "Form1.frx":0676
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   240
      OLEDropMode     =   1  'Manual
      Picture         =   "Form1.frx":07C0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":090A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nitro Tinylist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000CC99&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   270
      Width           =   1080
   End
End
Attribute VB_Name = "tinywindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim D As Boolean
Dim W As Boolean
Dim M
Dim SongCounter As Integer
Dim Z As Boolean
Dim S As Boolean
Dim A As Integer


Private Sub btn_Click(Index As Integer)
Select Case Index
Case 0 'about
frmMain.Show
Case 1 'play
PlayMP3
ResumeMP3
Case 2 'pause
PauseMP3
Case 3 'stop
StopMP3
Case 4 'rw
SongCounter = SongCounter - 1
If SongCounter = -1 Then SongCounter = 0
fileE.Text = M(SongCounter)
Case 5 'ff
If SongCounter = UBound(M) Then SongCounter = 0
fileE.Text = M(SongCounter)
SongCounter = SongCounter + 1
Case 6 'close
Unload Me
Case 7 'togle small/tinylist
If Me.Height = 2340 Then
Me.Height = 240
Else
Me.Height = 2340
End If
End Select
End Sub

Private Sub btn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
D = True
End Sub

Private Sub Form_Load()
Me.Top = Screen.Height
Me.Left = Screen.Width - Me.Width
If App.PrevInstance Then
If Not Command = "" Then SendMessage CLng(Val(GetSetting(App.Title, "ActiveWindow", "Handle"))), WM_SETTEXT, 0, ByVal CStr(Command)
End
Else
If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then SetDefaultDevice "MPEGVideo", "mciqtz.drv"
fileE.Text = Command
Z = True
S = False
End If
SaveSetting App.Title, "ActiveWindow", "Handle", Str(fileE.hWnd)
With nid
.cbSize = Len(nid)
.hWnd = Me.hWnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
A = 6
.hIcon = btn(0).Picture
.szTip = "http://redib.no-ip.com" & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long
If Me.ScaleMode = vbPixels Then
msg = X
Else
msg = X / Screen.TwipsPerPixelX
End If
Select Case msg
Case WM_RBUTTONUP
up.Enabled = False
Unload Me
End Select
up.Enabled = True
D = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
D = True
End Sub

Private Sub listy_DblClick()
SongCounter = listy.ListIndex
fileE.Text = M(SongCounter)
nexo.Enabled = True
nexo_Timer
End Sub

Private Sub listy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
D = True
End Sub

Private Sub up_Timer()
If D = True Then
D = False
If Me.Top < Screen.Height - (Me.Height + 400) Then: Exit Sub
Me.Top = Me.Top - 90
Else
If Me.Top > Screen.Height Then: up.Enabled = False: Exit Sub
Me.Top = Me.Top + 90
End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CloseAll
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub fileE_Change()
On Error Resume Next
If LCase(Right(fileE, 3)) = "mp3" Then
StopMP3
CloseMP3
OpenMP3 fileE
DoEvents
PlayMP3
alpha = Split(fileE, "\")
nid.szTip = Left(alpha(UBound(alpha)), Len(alpha(UBound(alpha))) - 4)
Shell_NotifyIcon NIM_MODIFY, nid
ElseIf LCase(Right(fileE, 3)) = "m3u" Then
For i = 0 To listy.ListCount
listy.RemoveItem i
Next i
SongCounter = 0
LoadList fileE
fileE.Text = M(SongCounter)
nexo.Enabled = True
nexo_Timer
End If
End Sub

Private Sub LoadList(liS As String)
Dim Data As String
o = Split(liS, "\")
o(UBound(o)) = ""
wer = Join(o, "\")
Open liS For Binary As #1
Data = Space(LOF(1))
Get #1, , Data
Close #1
X = Split(Data, vbCrLf)
For i = 0 To UBound(X) - 1
If Not Left(X(i), 1) = "#" Then
    If Mid(X(i), 2, 2) = ":\" Then
    Lista = Lista & X(i) & "?"
    Else
    Lista = Lista & wer & X(i) & "?"
    End If
End If
Next i
M = Split(Left(Lista, Len(Lista) - 1), "?")
For i = 0 To UBound(M)
phase = Split(M(i), "\")
RT = phase(UBound(phase))
listy.AddItem Left(RT, Len(RT) - 4), i
Next i
End Sub

Private Sub nexo_Timer()
  If AreMultimediaAtEnd(AliasName) = True Then
  If SongCounter = UBound(M) Then SongCounter = 0
  fileE.Text = M(SongCounter)
  SongCounter = SongCounter + 1
  End If
End Sub

Public Sub CloseMP3()
CloseMultimedia AliasName
End Sub

Public Sub OpenMP3(FileName As String)
  Dim typeDevice As String
  Dim Result As String
  typeDevice = "MPEGVideo"
  Result = OpenMultimedia(frmMain.hWnd, AliasName, FileName, typeDevice)
End Sub

Public Sub PauseMP3()
PauseMultimedia AliasName
End Sub

Public Sub PlayMP3()
PlayMultimedia AliasName, 0, 0
End Sub

Public Sub ResumeMP3()
ResumeMultimedia AliasName
End Sub

Public Sub StopMP3()
StopMultimedia AliasName
End Sub
