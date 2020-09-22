VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Final Fantasy RPG"
   ClientHeight    =   9765
   ClientLeft      =   4095
   ClientTop       =   2310
   ClientWidth     =   9765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   9765
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer timCheck 
      Left            =   7680
      Top             =   4680
   End
   Begin VB.Timer timwinlose 
      Left            =   7680
      Top             =   4200
   End
   Begin VB.Timer timhand 
      Left            =   7200
      Top             =   4200
   End
   Begin VB.Timer music2 
      Left            =   7680
      Top             =   5160
   End
   Begin VB.Timer music1 
      Left            =   7680
      Top             =   5640
   End
   Begin VB.Timer timdamageset 
      Interval        =   1
      Left            =   7200
      Top             =   5640
   End
   Begin VB.Timer timMon4 
      Left            =   3000
      Top             =   8160
   End
   Begin VB.Timer timMon3 
      Left            =   3000
      Top             =   7440
   End
   Begin VB.Timer timMon2 
      Left            =   1560
      Top             =   8160
   End
   Begin VB.Timer timallyattacks 
      Left            =   7200
      Top             =   4680
   End
   Begin VB.Timer timUserAttacks 
      Left            =   7200
      Top             =   5160
   End
   Begin VB.Frame fraAlly 
      Caption         =   "Ally's Information"
      Height          =   1935
      Left            =   7200
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
      Begin VB.Label lbldead 
         Alignment       =   1  'Right Justify
         Caption         =   "3"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblAllycheck 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label infoCase 
         Caption         =   "Exp:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   62
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblAllyAttack 
         Caption         =   "74"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblAexp 
         Caption         =   "100"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblallymp 
         Caption         =   "10"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblAllyhp 
         Caption         =   "100"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame informationUser 
      Caption         =   "User's Information"
      Height          =   2055
      Left            =   7200
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      Begin VB.Label lblUsercheck 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblyourdead 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblgroupgold 
         Caption         =   "400"
         Height          =   255
         Left            =   600
         TabIndex        =   63
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label infoCase 
         Caption         =   "Exp:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   61
         Top             =   960
         Width           =   495
      End
      Begin VB.Label infoCase 
         Caption         =   "Gold:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   60
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblUserAttack 
         Caption         =   "59"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblUexp 
         Caption         =   "500"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblmp 
         Caption         =   "30"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblhp 
         Caption         =   "200"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraMonsters 
      Caption         =   "Monsters:"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   7200
      Width           =   9735
      Begin VB.Timer timMon1 
         Left            =   1560
         Top             =   240
      End
      Begin VB.OptionButton optnone 
         Caption         =   "None"
         Height          =   255
         Left            =   4080
         TabIndex        =   80
         Top             =   1920
         Width           =   1935
      End
      Begin VB.OptionButton optzally 
         Caption         =   "Ally"
         Height          =   255
         Left            =   4080
         TabIndex        =   79
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton optzuser 
         Caption         =   "User"
         Height          =   255
         Left            =   4080
         TabIndex        =   78
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optmon2 
         Caption         =   "Monster 2"
         Height          =   255
         Left            =   4080
         TabIndex        =   74
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optmon4 
         Caption         =   "Monster 4"
         Height          =   255
         Left            =   4080
         TabIndex        =   32
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optmon3 
         Caption         =   "Monster 3"
         Height          =   255
         Left            =   4080
         TabIndex        =   31
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optmon1 
         Caption         =   "Monster 1"
         Height          =   255
         Left            =   4080
         TabIndex        =   30
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label lbldmagecount 
         Caption         =   "0"
         Height          =   255
         Left            =   4080
         TabIndex        =   88
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblmon2mp 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblmonAttack4 
         Caption         =   "0"
         Height          =   255
         Left            =   2760
         TabIndex        =   84
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblmonAttack3 
         Caption         =   "0"
         Height          =   255
         Left            =   2760
         TabIndex        =   83
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblmonAttack2 
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   82
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblmonAttack1 
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   81
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblmon2AC 
         Caption         =   "0"
         Height          =   255
         Left            =   5400
         TabIndex        =   77
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblmonMC2 
         Caption         =   "0"
         Height          =   255
         Left            =   6600
         TabIndex        =   76
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblmon2check 
         Caption         =   "0"
         Height          =   255
         Left            =   7680
         TabIndex        =   75
         Top             =   720
         Width           =   375
      End
      Begin VB.Label infoCase 
         Caption         =   "Gold:"
         Height          =   255
         Index           =   3
         Left            =   8280
         TabIndex        =   54
         Top             =   720
         Width           =   495
      End
      Begin VB.Label infoCase 
         Caption         =   "Exp:"
         Height          =   255
         Index           =   2
         Left            =   8280
         TabIndex        =   53
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblgold 
         Caption         =   "0"
         Height          =   255
         Left            =   8760
         TabIndex        =   52
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblexp 
         Caption         =   "0"
         Height          =   255
         Left            =   8640
         TabIndex        =   51
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblalldead 
         Caption         =   "0"
         Height          =   255
         Left            =   6600
         TabIndex        =   50
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblmon4check 
         Caption         =   "0"
         Height          =   255
         Left            =   7680
         TabIndex        =   49
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblmon3check 
         Caption         =   "0"
         Height          =   255
         Left            =   7680
         TabIndex        =   48
         Top             =   960
         Width           =   375
      End
      Begin VB.Label infoCase 
         Caption         =   "Reset Case"
         Height          =   255
         Index           =   1
         Left            =   7320
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblmon1check 
         Caption         =   "0"
         Height          =   255
         Left            =   7680
         TabIndex        =   46
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblmonMC4 
         Caption         =   "0"
         Height          =   255
         Left            =   6600
         TabIndex        =   41
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblmon4AC 
         Caption         =   "0"
         Height          =   255
         Left            =   5400
         TabIndex        =   40
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblmonMC3 
         Caption         =   "0"
         Height          =   255
         Left            =   6600
         TabIndex        =   39
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblmon3AC 
         Caption         =   "0"
         Height          =   255
         Left            =   5400
         TabIndex        =   38
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblmonMC 
         Caption         =   "0"
         Height          =   255
         Left            =   6600
         TabIndex        =   37
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblmon1AC 
         Caption         =   "0"
         Height          =   255
         Left            =   5400
         TabIndex        =   36
         Top             =   480
         Width           =   255
      End
      Begin VB.Label infoCase 
         Caption         =   "Monster Case:"
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label infoAC 
         Caption         =   "Attack Case:"
         Height          =   255
         Left            =   5160
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
      Begin VB.Label infoAttack 
         Caption         =   "Attacking:"
         Height          =   255
         Left            =   4200
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblmon2hp 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblmon2name 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblmon4mp 
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblmon4hp 
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblmon4name 
         Caption         =   "Name:"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblmon3mp 
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblmon3hp 
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblmon3name 
         Caption         =   "Name:"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblmonmp 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblmonhp 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblmonName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   8160
      TabIndex        =   89
      Top             =   4200
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label lblallydam 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   71
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lbluserdam 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   70
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image imghandBattle 
      Height          =   480
      Left            =   6480
      Picture         =   "frmMain.frx":08CA
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label mblallympMax 
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   69
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label infoCase 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   6000
      TabIndex        =   68
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label mblallymp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   67
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label mbluserMaxmp 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   66
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label infoCase 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   6000
      TabIndex        =   65
      Top             =   5640
      Width           =   135
   End
   Begin VB.Label mbluserMP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   64
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image imgallyhand 
      Height          =   480
      Left            =   6600
      Picture         =   "frmMain.frx":0C84
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imguserhand 
      Height          =   480
      Left            =   6600
      Picture         =   "frmMain.frx":154E
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgmonhand4 
      Height          =   480
      Left            =   2160
      Picture         =   "frmMain.frx":1E18
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgmonhand3 
      Height          =   480
      Left            =   2160
      Picture         =   "frmMain.frx":26E2
      Top             =   6120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgmonhand2 
      Height          =   480
      Left            =   2160
      Picture         =   "frmMain.frx":2FAC
      Top             =   5760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgmonhand1 
      Height          =   480
      Left            =   2160
      Picture         =   "frmMain.frx":3876
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label infoCase 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   59
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label infoCase 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   58
      Top             =   5640
      Width           =   135
   End
   Begin VB.Label lblAMaxlife 
      BackStyle       =   0  'Transparent
      Caption         =   "150"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   57
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label lblUMaxLife 
      BackStyle       =   0  'Transparent
      Caption         =   "300"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   56
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lbluserlife 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "300"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   55
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lblmon4dam 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   45
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label lblmon3dam 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   44
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblmon2dam 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   43
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label lblmon1dam 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   42
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgmon4 
      Height          =   720
      Left            =   240
      Picture         =   "frmMain.frx":4140
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image imgmon3 
      Height          =   720
      Left            =   840
      Picture         =   "frmMain.frx":46AA
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image imgmon2 
      Height          =   720
      Left            =   1080
      Picture         =   "frmMain.frx":4C14
      Top             =   4200
      Width           =   720
   End
   Begin VB.Image imgmon1 
      Height          =   720
      Left            =   1680
      Picture         =   "frmMain.frx":51C5
      Top             =   3480
      Width           =   720
   End
   Begin VB.Image imgally 
      Height          =   360
      Left            =   5640
      Picture         =   "frmMain.frx":5776
      Top             =   4080
      Width           =   240
   End
   Begin VB.Image imguser 
      Height          =   360
      Left            =   5520
      Picture         =   "frmMain.frx":5B53
      Top             =   3840
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   135
      Index           =   1
      Left            =   4920
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Final Fantasy RPG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label lblmon4N 
      BackStyle       =   0  'Transparent
      Caption         =   "Monster #4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label lblmon3N 
      BackStyle       =   0  'Transparent
      Caption         =   "Monster #3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label lblmon2N 
      BackStyle       =   0  'Transparent
      Caption         =   "Monster #2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lblmon1N 
      BackStyle       =   0  'Transparent
      Caption         =   "Monster #1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Shape shpAA 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      FillStyle       =   4  'Upward Diagonal
      Height          =   135
      Left            =   4920
      Top             =   6000
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   135
      Index           =   0
      Left            =   4920
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Shape shpUA 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   135
      Left            =   4920
      Top             =   5520
      Width           =   15
   End
   Begin VB.Label lblallylife 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "150"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label lblallyname 
      BackStyle       =   0  'Transparent
      Caption         =   "Soldier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "??????"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Image imgmenu 
      Height          =   1845
      Left            =   0
      Picture         =   "frmMain.frx":5F30
      Top             =   5280
      Width           =   7155
   End
   Begin VB.Image imgstatus 
      Height          =   930
      Left            =   0
      Picture         =   "frmMain.frx":30F7A
      Top             =   0
      Width           =   7200
   End
   Begin VB.Image imgfightbg 
      Height          =   4440
      Left            =   0
      Picture         =   "frmMain.frx":46C7C
      Top             =   840
      Width           =   7200
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Game"
         Enabled         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSound 
      Caption         =   "&Sound"
      Begin VB.Menu mnuOn 
         Caption         =   "On"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOff 
         Caption         =   "Off"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mark - $uper-$tar
'Final Fantasy RPG
'Main Battle
Option Explicit
Dim intgold As Integer
Dim intexp As Integer
Dim intrand1 As Integer 'Attack Monster 1
Dim intrand2 As Integer
Dim intlost As Integer 'MSGBOX REPLY
'Music
    Dim PLAYLOOP As Integer

'[What Things Do:]
    'imgstatus = The Status Picture
    'imgfightbg = The Battle Back Ground
    'imgMenu = The Menu
    'shpUA = Shape of Attack Bar
    'shpAA = Shape of Ally's Attack Bar
    'cmbUAttacks = User Attacks
    'cmbAattacks = Ally ATtacks
    'timUserAttacks = User Attack Timer
    'AC# = Attacks, Wacth out for Errors in these
    'MC# = Monster ID number

Private Sub Form_Load()
        frmBegin.Visible = False
        frmBegin.music1.Enabled = False
        frmBegin.music2.Enabled = False
        frmBegin.MMControl1.Command = "Stop"
'This sets Form to play "game"
        frmAttack.Visible = False
        frmMain.Width = 7155
        frmMain.Height = 7770
        lblmon1dam.BackStyle = 0
        lblmon2dam.BackStyle = 0
        lblmon3dam.BackStyle = 0
        lblmon4dam.BackStyle = 0
'Timers
        timUserAttacks.Interval = 60 'User Attacks Tim.
        timallyattacks.Interval = 60 'Ally Attacks Tim.
        timCheck.Interval = 60 'Clears Data
        timwinlose.Interval = 60
        timhand.Interval = 60
        timdamageset.Interval = 99 'Damage Set
'monsters:
    timMon1.Interval = 60
    timMon2.Interval = 60
    timMon3.Interval = 60
    timMon4.Interval = 60
'Monster
        randommonster 'Monster
'Text
        lblstatus.Caption = "Final Fantasy RPG"
        lblmon1dam.Caption = ""
        mbluserMP.Caption = "50"
        mblallymp.Caption = "30"
'Picture:
        frmMain.imguser.Picture = frmPic.imguser(2).Picture
        frmMain.imgally.Picture = frmPic.Image2(2).Picture
'Moves Guy
        frmMain.imguser.Left = 5520
        frmMain.imgally.Left = 5640
'Music
    MMControl1.FileName = App.Path + "\Battle.mid"
    MMControl1.Command = "Open"
    MMControl1.Command = "Play"
    PLAYLOOP = True
End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub lblallyname_Click()
optzally.Value = True
End Sub

Private Sub lblmon1N_Click()
optmon1.Value = True 'Attack this target
End Sub

Private Sub lblmon2N_Click()
optmon2.Value = True 'Attack this target
End Sub

Private Sub lblmon3N_Click()
optmon3.Value = True 'Attack this target
End Sub

Private Sub lblmon4N_Click()
optmon4.Value = True 'Attack this target
End Sub

Private Sub mnuEnd_Click()
End
End Sub

Private Sub lblUserName_Click()
optzuser.Value = True
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuNew_Click()
'Timers
    timUserAttacks.Interval = 60 'User Attacks Tim.
    timallyattacks.Interval = 60 'Ally Attacks Tim.
    timCheck.Interval = 60 'Clears Data
    timwinlose.Interval = 60
    timhand.Interval = 60
    timdamageset.Interval = 99 'Damage Set
'monsters:
    timMon1.Interval = 60
    timMon2.Interval = 60
    timMon3.Interval = 60
    timMon4.Interval = 60
'Monster
        randommonster 'Monster
'Text
        'Status
        lblstatus.Caption = "Final Fantasy RPG"
        'Magic
        lblmon1dam.Caption = ""
        mbluserMP.Caption = "50"
        mblallymp.Caption = "30"
        'Life
        lbluserlife.Caption = "300"
        lblallylife.Caption = "150"
        lblAMaxlife.Caption = "150"
        lblUMaxLife.Caption = "300"
        'Attack
        lblUserAttack.Caption = "0"
        lblAllyAttack.Caption = "0"
        shpAA.Width = "15"
        shpUA.Width = "15"
        'Monster Attack
        lblmonAttack1.Caption = "2"
        lblmonAttack2.Caption = "-7"
        lblmonAttack3.Caption = "-2"
        lblmonAttack4.Caption = "0"
        
'Picture:
        frmMain.imguser.Picture = frmPic.imguser(2).Picture
        frmMain.imgally.Picture = frmPic.Image2(2).Picture
'Moves Guy
        frmMain.imguser.Left = 5520
        frmMain.imgally.Left = 5640
    'Items
        frmAttack.cmbitems.RemoveItem frmAttack.cmbitems.ListIndex
        frmAttack.cmbitems.AddItem ("Potion")
        frmAttack.cmbitems.AddItem ("Potion")
        frmAttack.cmbitems.AddItem ("Potion")
        frmAttack.cmbitems.AddItem ("Tincture")
        frmAttack.cmbitems.ListIndex = 0

'Music
    MMControl1.FileName = App.Path + "\Battle.mid"
    MMControl1.Command = "Open"
    MMControl1.Command = "Play"
    PLAYLOOP = True
End Sub

Private Sub mnuOff_Click()
'Music = Off
mnuOn.Checked = False
mnuOff.Checked = True
MMControl1.Command = "Stop"
End Sub

Private Sub mnuOn_Click()
'Music = On
mnuOn.Checked = True
mnuOff.Checked = False
MMControl1.Command = "Play"
End Sub

Private Sub timallyattacks_Timer()
'This  "if" resets:
If lblAllyAttack.Caption = 15 Or lblAllyAttack.Caption = 30 _
Or lblAllyAttack.Caption = 45 Or lblAllyAttack.Caption = 60 Then
'Set Reset Damage
    imgally.Picture = frmPic.Image2(2).Picture
'Moves Guy
    frmMain.imgally.Left = 5640
End If
If lblAllyAttack.Caption = 75 Then
    'Sets Buttons:
    frmAttack.cmbitems.Enabled = True
    frmAttack.cmbUAttacks.Enabled = True
    frmAttack.cmbUMagic.Enabled = True
    frmAttack.cmbAattacks.Enabled = True
    frmAttack.cmbAmagic.Enabled = True
    frmAttack.cmbUAttacks.Visible = False
    frmAttack.cmbUMagic.Visible = False
    frmAttack.cmbAattacks.Visible = True
    frmAttack.cmbAmagic.Visible = True
    'Changes Picture
    imgally.Picture = frmPic.Image2(0)
    'Makes Attack Visibe
    frmAttack.Visible = True
    'Timers
    timUserAttacks.Enabled = False
Else
    'Adds to Counter
    lblAllyAttack.Caption = lblAllyAttack.Caption + 1
    imgallyhand.Visible = False
    shpAA.Width = shpAA.Width + 19
    'Timers
    timUserAttacks.Enabled = True
End If
End Sub


Private Sub timcheck_Timer() 'Clears Data
'No Attack
If optnone.Value = True Then
    imgmonhand1.Visible = False
    imgmonhand2.Visible = False
    imgmonhand3.Visible = False
    imgmonhand4.Visible = False
    imghandBattle.Visible = False
End If

'Dead:
'Monster 1: (Dead)
If lblmonhp.Caption < 0 Then
    lblmon1N.Visible = False 'Disabes
    imgmon1.Visible = False 'Disbles
'Stops Attacks:
    lblmonhp.Caption = 0
    optmon1.Visible = False
'Makes it "dead"
    lblalldead.Caption = lblalldead.Caption - 1
    timMon1.Enabled = False
    optnone.SetFocus
End If

'Monster 2: (Dead)
If lblmon2hp.Caption < 0 Then
    lblmon2N.Visible = False 'Disabes
    imgmon2.Visible = False 'Disbles
'Stops Attacks:
    optmon2.Visible = False
    lblmon2hp.Caption = 0
'Makes it "dead"
    lblalldead.Caption = lblalldead.Caption - 1
    timMon2.Enabled = False
    optnone.SetFocus
End If
'Monster 3: (Dead)
If lblmon3hp.Caption < 0 Then
    lblmon3N.Visible = False 'Disabes
    imgmon3.Visible = False 'Disbles
'Stops Attacks:
    optmon3.Visible = False
    lblmon3hp.Caption = 0
'Makes it "dead"
    lblalldead.Caption = lblalldead.Caption - 1
    timMon3.Enabled = False
    optnone.SetFocus
End If
'Monster 4: (Dead)
If lblmon4hp.Caption < 0 Then
    lblmon4N.Visible = False 'Disabes
    imgmon4.Visible = False 'Disbles
'Stops Attacks:
    optmon4.Visible = False
    lblmon4hp.Caption = 0
'Makes it "dead"
    timMon4.Enabled = False
    lblalldead.Caption = lblalldead.Caption - 1
        optnone.SetFocus
End If
'Damage:
'Damage Data For Monster 1
    If lblmon1check.Caption = 15 Then
        lblmon1dam.Caption = "" 'Resets Damage
        lblmon1check.Caption = 0 'Resets Timer
    Else
        lblmon1check.Caption = lblmon1check.Caption + 1
    End If
'Damage Data For Monster 2
    If lblmon2check.Caption = 15 Then
        lblmon2dam.Caption = ""
        lblmon2check.Caption = 0
    Else
        lblmon2check.Caption = lblmon2check.Caption + 1
    End If
'Damage Data For Monster 3
    If lblmon3check.Caption = 15 Then
        lblmon3dam.Caption = ""
        lblmon3check.Caption = 0
    Else
        lblmon3check.Caption = lblmon3check.Caption + 1
    End If
'Damage Data For Monster 4
    If lblmon4check.Caption = 15 Then
        lblmon4dam.Caption = ""
        lblmon4check.Caption = 0
    Else
        lblmon4check.Caption = lblmon4check.Caption + 1
    End If
'Damage Data for User
If lblUsercheck.Caption = 15 Then
    lbluserdam.Caption = ""
    lblUsercheck.Caption = 0
    lblstatus.Caption = "" 'Clears Status
Else
lblUsercheck.Caption = lblUsercheck.Caption + 1
End If
'Wound:
If lbluserlife.Caption >= 1 And lbluserlife.Caption <= 45 Then
'Picture
    imguser.Picture = frmPic.imguser(4).Picture
End If
'Death:
If lbluserlife.Caption <= 0 Then
'Stops Attacks:
    imguser.Picture = frmPic.imguser(5).Picture
' Makes it "dead"
    lblyourdead.Caption = lblyourdead.Caption - 1
End If
    
'Damage Data for ally
If lblAllycheck.Caption = 15 Then
    lblallydam.Caption = ""
    lblAllycheck.Caption = 0
    lblstatus.Caption = "" 'Clears Status
Else
lblAllycheck.Caption = lblAllycheck.Caption + 1
End If
'Wound:
If lblallylife.Caption >= 1 And lblallylife.Caption <= 45 Then
'Picture
    imgally.Left = 5760
    imgally.Top = 4440
    imgally.Picture = frmPic.Image2(5).Picture
End If
'Death:
If lblallylife.Caption <= 0 Then
lblallylife.Caption = 0
lbldead.Caption = 2
'Stops Attacks:
timallyattacks.Enabled = False
shpAA.Width = 15
lblAllyAttack.Caption = 2
'Pics
    imgally.Picture = frmPic.Image2(3).Picture

End If
End Sub

Private Sub timdamageset_Timer()
'Stop Flicker?
    If lbldmagecount.Caption = 5 Then
    timdamageset.Enabled = False
    Else
    timdamageset.Enabled = True
'Monster 1:
    lblmon1dam.Left = imgmon1.Left + (imgmon1.Width / 2)
    lblmon1dam.Top = imgmon1.Top + (imgmon1.Height / 2)
'Monster 2:
    lblmon2dam.Left = imgmon2.Left + (imgmon2.Width / 2)
    lblmon2dam.Top = imgmon2.Top + (imgmon2.Height / 2)
'Monster 3:
    lblmon3dam.Left = imgmon3.Left + (imgmon3.Width / 2)
    lblmon3dam.Top = imgmon3.Top + (imgmon3.Height / 2)
'Monster 4:
    lblmon4dam.Left = imgmon4.Left + (imgmon4.Width / 2)
    lblmon4dam.Top = imgmon4.Top + (imgmon4.Height / 2)
'And Counter
lbldmagecount.Caption = lbldmagecount.Caption + 1
    End If

End Sub


Private Sub timhand_Timer()
If optzuser.Value = True Then
'Hand:
imghandBattle.Left = imguser.Left + imguser.Width
imghandBattle.Top = imguser.Top + (imguser.Height / 2)
'Pics:
imgmonhand1.Visible = False
imgmonhand2.Visible = False
imgmonhand3.Visible = False
imgmonhand4.Visible = False
End If

If optzally.Value = True Then
'Hand:
imghandBattle.Left = imgally.Left + imgally.Width
imghandBattle.Top = imgally.Top + (imgally.Height / 2)
'Pics:
imgmonhand1.Visible = False
imgmonhand2.Visible = False
imgmonhand3.Visible = False
imgmonhand4.Visible = False
End If

If optmon1.Value = True Then
imgmonhand1.Visible = True
'Hand:
imghandBattle.Visible = True
imghandBattle.Left = imgmon1.Left + imgmon1.Width
imghandBattle.Top = imgmon1.Top + (imgmon1.Height / 2)
'Pictures
imgmonhand2.Visible = False
imgmonhand3.Visible = False
imgmonhand4.Visible = False
End If

If optmon2.Value = True Then
imgmonhand2.Visible = True
'Hand:
imghandBattle.Visible = True
imghandBattle.Left = imgmon2.Left + imgmon2.Width
imghandBattle.Top = imgmon2.Top + (imgmon2.Height / 2)
'Pictures
imgmonhand1.Visible = False
imgmonhand3.Visible = False
imgmonhand4.Visible = False
End If

If optmon3.Value = True Then
imgmonhand3.Visible = True
'Hand:
imghandBattle.Visible = True
imghandBattle.Left = imgmon3.Left + imgmon3.Width
imghandBattle.Top = imgmon3.Top + (imgmon3.Height / 2)
'Pictures
imghandBattle.Visible = True
imgmonhand2.Visible = False
imgmonhand1.Visible = False
imgmonhand4.Visible = False
End If

If optmon4.Value = True Then
imgmonhand4.Visible = True
'Hand:
imghandBattle.Visible = True
imghandBattle.Left = imgmon4.Left + imgmon4.Width
imghandBattle.Top = imgmon4.Top + (imgmon4.Height / 2)
'Pictures
imgmonhand2.Visible = False
imgmonhand3.Visible = False
imgmonhand1.Visible = False
End If
End Sub

Private Sub timMon1_Timer()
If lblmonAttack1.Caption = 65 Then
'Attack
  Mon1Attack
Else
'Adds to Counter
    lblmonAttack1.Caption = lblmonAttack1.Caption + 1
End If
End Sub

Private Sub timMon2_Timer()
If lblmonAttack2.Caption = 60 Then
'Attack
    Mon2Attack
Else
'Adds to Counter
    lblmonAttack2.Caption = lblmonAttack2.Caption + 1
End If
End Sub

Private Sub timMon3_Timer()
If lblmonAttack3.Caption = 75 Then
'Attack
    Mon3Attack
Else
'Adds to Counter
    lblmonAttack3.Caption = lblmonAttack3.Caption + 1
End If
End Sub

Private Sub timMon4_Timer()
If lblmonAttack4.Caption = 60 Then
'Attack
    Mon4Attack
Else
'Adds to Counter
    lblmonAttack4.Caption = lblmonAttack4.Caption + 1
End If
End Sub

Private Sub timUserAttacks_Timer()
'This  "if" resets:
If lblUserAttack.Caption = 15 Or lblUserAttack.Caption = 30 _
Or lblUserAttack.Caption = 45 Then
'Set Reset Damage
    imguser.Picture = frmPic.imguser(2).Picture
'Moves Guy
    frmMain.imguser.Left = 5520
End If
If lblUserAttack.Caption = 60 Then
'Sets Buttons:
    frmAttack.cmbitems.Enabled = True
    frmAttack.cmbUAttacks.Enabled = True
    frmAttack.cmbUMagic.Enabled = True
    frmAttack.cmbAattacks.Enabled = True
    frmAttack.cmbAmagic.Enabled = True
    frmAttack.cmbUAttacks.Visible = True
    frmAttack.cmbUMagic.Visible = True
    frmAttack.cmbAattacks.Visible = False
    frmAttack.cmbAmagic.Visible = False
'Changes Picture
    imguser.Picture = frmPic.imguser(1)
    imguserhand.Visible = True
'Makes Attack Visible
    frmAttack.Visible = True
'Stop Others
    timallyattacks.Enabled = False
Else
'Adds to Counter
    lblUserAttack.Caption = lblUserAttack.Caption + 1
    shpUA.Width = shpUA.Width + 24
    imguserhand.Visible = False
'Stop Others
    timallyattacks.Enabled = True
End If
End Sub

Private Sub timwinlose_Timer()
'To Win
If lblalldead.Caption <= 0 Then
'Win Screen
    intexp = lblexp.Caption
    intgold = lblgold.Caption
    lblUexp.Caption = lblUexp.Caption + intexp
    lblAexp.Caption = lblAexp.Caption + intexp
    lblgroupgold.Caption = lblgroupgold.Caption + intgold
'Shows Win Screen
    frmBattleWin.Visible = True
'Turns Of Timer
    timwinlose.Enabled = False
'Changes invisible
    lblmon1dam.Visible = False
    lblmon2dam.Visible = False
    lblmon3dam.Visible = False
    lblmon4dam.Visible = False
End If
If lblyourdead.Caption <= 0 Then
    lblstatus.Caption = "You have died! Game Over"
    timallyattacks.Enabled = False
    timUserAttacks.Enabled = False
    timMon1.Enabled = False
    timMon2.Enabled = False
    timMon3.Enabled = False
    timMon4.Enabled = False
    timhand.Enabled = False
    timCheck.Enabled = False
intlost = MsgBox("You Have Lost, Do you want to play again?", vbCritical + vbYesNo, "Game Over")
Select Case intlost
Case vbYes
    'Resets:
    lbluserdam.Caption = ""
    lblallydam.Caption = ""
    lblmon1dam.Caption = ""
    lblmon2dam.Caption = ""
    lblmon3dam.Caption = ""
    lblmon4dam.Caption = ""
    'Attacks
    lblUserAttack.Caption = 0
    shpUA.Width = 15
    lblAllyAttack.Caption = 0
    shpAA.Width = 15
    'Deaths
    lblyourdead.Caption = 1
    lbldead.Caption = 3
    'Text
    lbluserlife.Caption = 300
    lblallylife.Caption = 150
    mbluserMP.Caption = 50
    mblallymp.Caption = 30
    lblstatus.Caption = "Final Fantasy RPG"
    'Reset Timers
    timallyattacks.Enabled = True
    timUserAttacks.Enabled = True
    timMon1.Enabled = True
    timMon2.Enabled = True
    timMon3.Enabled = True
    timMon4.Enabled = True
    timhand.Enabled = True
    timCheck.Enabled = True
    'Random Monster:
    randommonster
    'Picture:
        frmMain.imguser.Picture = frmPic.imguser(2).Picture
        frmMain.imgally.Picture = frmPic.Image2(2).Picture
    'Moves Guy
        frmMain.imguser.Left = 5520
        frmMain.imgally.Left = 5640
Case vbNo
    End
Case Else
End Select
End If
End Sub
'*********************[Music]********************'
Private Sub music1_Timer()
music2.Enabled = True
music1.Enabled = False
End Sub

Private Sub music2_Timer()
'Music
MMControl1.Command = "Stop"
MMControl1.FileName = App.Path + "\Battle.mid"
MMControl1.DeviceType = "Sequencer"
MMControl1.Command = "Seek"
MMControl1.Command = "Open"
MMControl1.Command = "Play"
PLAYLOOP = True
music1.Enabled = True
music2.Enabled = False
End Sub

