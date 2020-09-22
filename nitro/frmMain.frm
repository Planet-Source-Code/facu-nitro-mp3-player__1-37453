VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Redib Warfare's Nitro! Mp3 Player"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   1200
      Picture         =   "frmMain.frx":091A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   24
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   1440
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":1234
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   1680
      Picture         =   "frmMain.frx":137E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   1920
      Picture         =   "frmMain.frx":14C8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   4
      Left            =   2160
      Picture         =   "frmMain.frx":1612
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   20
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   5
      Left            =   2400
      Picture         =   "frmMain.frx":175C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox btn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   6
      Left            =   2640
      Picture         =   "frmMain.frx":18A6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Picture2 
      Height          =   315
      Left            =   360
      ScaleHeight     =   255
      ScaleWidth      =   1575
      TabIndex        =   4
      Top             =   2280
      Width           =   1635
      Begin VB.CommandButton hid 
         BackColor       =   &H0000CC99&
         Caption         =   "Ok, Now Close!"
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2025
      Left            =   0
      ScaleHeight     =   1965
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   240
      Width           =   315
      Begin VB.CommandButton cmdRegister 
         BackColor       =   &H0000CC99&
         Caption         =   "&Associate"
         Height          =   1960
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox wesc 
      BackColor       =   &H00FFFFFF&
      Height          =   2025
      Left            =   360
      ScaleHeight     =   1965
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   240
      Width           =   3375
      Begin VB.PictureBox tosc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6015
         Left            =   0
         ScaleHeight     =   6015
         ScaleWidth      =   3375
         TabIndex        =   6
         Top             =   0
         Width           =   3375
         Begin VB.Image Image2 
            Height          =   240
            Left            =   120
            Picture         =   "frmMain.frx":19F0
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nitro! Mp3 Player"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00CC9900&
            Height          =   240
            Left            =   960
            TabIndex        =   17
            Top             =   120
            Width           =   1680
         End
         Begin VB.Label lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "What is Nitro! Mp3 Player?"
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   1905
         End
         Begin VB.Label lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nitro! Is the simpliest Mp3 Player Ever."
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   15
            Top             =   840
            Width           =   2790
         End
         Begin VB.Label lab 
            BackStyle       =   0  'Transparent
            Caption         =   "Why should I use Nitro! instead of any other Mp3 Player?"
            ForeColor       =   &H00808080&
            Height          =   435
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   2970
         End
         Begin VB.Label lab 
            BackStyle       =   0  'Transparent
            Caption         =   "Because it's for music lovers. There isn't even a window to distract you. All your attention goes directly to music."
            Height          =   675
            Index           =   3
            Left            =   240
            TabIndex        =   13
            Top             =   1800
            Width           =   2925
         End
         Begin VB.Label lab 
            BackStyle       =   0  'Transparent
            Caption         =   "Does this crapy program uses too much memory?"
            ForeColor       =   &H00808080&
            Height          =   435
            Index           =   4
            Left            =   120
            TabIndex        =   12
            Top             =   2640
            Width           =   2925
         End
         Begin VB.Label lab 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":230A
            Height          =   1035
            Index           =   5
            Left            =   240
            TabIndex        =   11
            Top             =   3120
            Width           =   2865
         End
         Begin VB.Label lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Do I have to pay anything?"
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   10
            Top             =   4440
            Width           =   1965
         End
         Begin VB.Label lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No, this software is absolutly free."
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   9
            Top             =   4680
            Width           =   2490
         End
         Begin VB.Label lab 
            BackStyle       =   0  'Transparent
            Caption         =   "Check out Redib Warfare's Web Site for More Quality Software!"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000CC99&
            Height          =   435
            Index           =   8
            Left            =   120
            TabIndex        =   8
            Top             =   5160
            Width           =   3135
         End
         Begin VB.Label page 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "http://redib.no-ip.com"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            MousePointer    =   10  'Up Arrow
            TabIndex        =   7
            Top             =   5640
            Width           =   3015
         End
      End
   End
   Begin VB.VScrollBar scl 
      Height          =   2025
      Left            =   3720
      Max             =   100
      TabIndex        =   0
      Top             =   240
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   1965
      Picture         =   "frmMain.frx":23AF
      Top             =   0
      Width           =   1905
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRegister_Click()
  Dim MyFileType As filetype
  Path = App.Path
  If Right$(Path, 1) <> "\" Then Path = Path & "\"
  Path = Path & App.EXEName & ".exe"
    'mp3
    MyFileType.ProperName = "MPlayer3"
    MyFileType.FullName = "MPlayer3 MP3 File"
    MyFileType.ContentType = "audio/mp3"
    MyFileType.extension = ".mp3"
    MyFileType.Commands.Captions.Add "Open"
    MyFileType.Commands.Commands.Add Chr$(34) & Path & Chr$(34) & " " & "%1"
    MyFileType.IconPath = Path
    MyFileType.IconIndex = 0
    CreateExtension MyFileType
    'm3u
    MyFileType.ProperName = "MPlayer3"
    MyFileType.FullName = "MPlayer3 Playlist"
    MyFileType.ContentType = "audio/m3u"
    MyFileType.extension = ".m3u"
    MyFileType.Commands.Captions.Add "Open"
    MyFileType.Commands.Commands.Add Chr$(34) & Path & Chr$(34) & " " & "%1"
    MyFileType.IconPath = Path
    MyFileType.IconIndex = 0
    CreateExtension MyFileType
End Sub

Private Sub hid_Click()
Unload Me
End Sub

Private Sub scl_Change()
tosc.Top = (scl.Value * (wesc.Height - tosc.Height)) / scl.Max
End Sub

Private Sub scl_Scroll()
tosc.Top = (scl.Value * (wesc.Height - tosc.Height)) / scl.Max
End Sub
