VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmBegin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Final Fantasy RPG"
   ClientHeight    =   6210
   ClientLeft      =   4455
   ClientTop       =   3390
   ClientWidth     =   7140
   Icon            =   "frmBegin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmBegin.frx":08CA
   MousePointer    =   99  'Custom
   Picture         =   "frmBegin.frx":1194
   ScaleHeight     =   6210
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimSelect 
      Interval        =   7
      Left            =   240
      Top             =   4080
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   3360
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer music2 
      Left            =   240
      Top             =   4560
   End
   Begin VB.Timer music1 
      Left            =   240
      Top             =   5040
   End
   Begin VB.Image img3 
      Height          =   480
      Left            =   7320
      Picture         =   "frmBegin.frx":935D6
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image img2 
      Height          =   480
      Left            =   7320
      Picture         =   "frmBegin.frx":93EA0
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image img1 
      Height          =   480
      Left            =   7320
      Picture         =   "frmBegin.frx":9476A
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   4600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Load Game"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2800
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   4000
      Width           =   1335
   End
   Begin VB.Label lblBattle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Battle"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3400
      Width           =   1335
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2200
      Width           =   1335
   End
   Begin VB.Label lblNAme 
      BackStyle       =   0  'Transparent
      Caption         =   "Final Fantasy RPG"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mark - $uper-$tar
'Final Fantasy RPG
'Begin
'Music
    Dim PLAYLOOP As Integer
    Dim intmouseicon As Integer
Private Sub Form_Load()
'Mouse
TimSelect.Interval = 60
'Music
MMControl1.FileName = App.Path + "\FF6EMPIR.MID"
MMControl1.Command = "Open"
MMControl1.DeviceType = "Sequencer"
MMControl1.Command = "Play"
PLAYLOOP = True
End Sub

Private Sub lblBattle_Click()
frmMain.Visible = True
End Sub

Private Sub lblEnd_Click()
End
End Sub

Private Sub lblInfo_Click()
frmInfo.Show vbModal, Me
End Sub

'*********************[Music]********************'
Private Sub music1_Timer()
music2.Enabled = True
music1.Enabled = False
End Sub

Private Sub music2_Timer()
'Music
MMControl1.Command = "Stop"
MMControl1.FileName = App.Path + "\FF6EMPIR.MID"
MMControl1.DeviceType = "Sequencer"
MMControl1.Command = "Seek"
MMControl1.Command = "Open"
MMControl1.Command = "Play"
PLAYLOOP = True
music1.Enabled = True
music2.Enabled = False
End Sub

Private Sub TimSelect_Timer()

intmouseicon = intmouseicon + 1

If intmouseicon = 1 Then
frmBegin.MouseIcon = img1.Picture
End If

If intmouseicon = 3 Then
frmBegin.MouseIcon = img2.Picture
End If

If intmouseicon = 5 Then
frmBegin.MouseIcon = img3.Picture
intmouseicon = -1
End If

End Sub
