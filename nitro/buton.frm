VERSION 5.00
Begin VB.Form buton 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleWidth      =   1680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "buton.frx":0000
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
      Index           =   1
      Left            =   240
      OLEDropMode     =   1  'Manual
      Picture         =   "buton.frx":091A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
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
      Picture         =   "buton.frx":0A64
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   720
      Picture         =   "buton.frx":0BAE
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
      Index           =   4
      Left            =   960
      Picture         =   "buton.frx":0CF8
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
      Index           =   5
      Left            =   1200
      Picture         =   "buton.frx":0E42
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
      Index           =   6
      Left            =   1440
      Picture         =   "buton.frx":0F8C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   0
      Width           =   240
   End
   Begin VB.Timer up 
      Interval        =   50
      Left            =   1680
      Top             =   0
   End
End
Attribute VB_Name = "buton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = Screen.Height
Me.Left = Screen.Width - Me.Width
End Sub

Private Sub up_Timer()
If C = True Then
C = False
If Me.Top < Screen.Height - Me.Height + 300 Then Exit Sub
Me.Top = Me.Top - 30
Else
If Me.Top > Screen.Height Then: up.Enabled = False: Exit Sub
Me.Top = Me.Top + 30
End If
End Sub

Private Sub btn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'C = True
MsgBox C
Select Case Index
Case 0 'about
Case 1 'play
Case 2 'pause
Case 3 'stop
Case 4 'rw
Case 5 'ff
Case 6 'close
End Select
End Sub
