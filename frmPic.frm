VERSION 5.00
Begin VB.Form frmPic 
   Caption         =   "Pictures"
   ClientHeight    =   5100
   ClientLeft      =   3045
   ClientTop       =   2835
   ClientWidth     =   9525
   Icon            =   "frmPic.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmPic.frx":08CA
   ScaleHeight     =   5100
   ScaleWidth      =   9525
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Ally Pictures:"
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ally Pictures:"
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Frame fraSeigfried 
      Caption         =   "Seigfried Pictures:"
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   3855
      Begin VB.Image imgSeigfried 
         Height          =   360
         Index           =   4
         Left            =   1560
         Picture         =   "frmPic.frx":09A9
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgSeigfried 
         Height          =   360
         Index           =   3
         Left            =   1200
         Picture         =   "frmPic.frx":0DA8
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgSeigfried 
         Height          =   360
         Index           =   2
         Left            =   840
         Picture         =   "frmPic.frx":119F
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgSeigfried 
         Height          =   360
         Index           =   1
         Left            =   480
         Picture         =   "frmPic.frx":1590
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgSeigfried 
         Height          =   360
         Index           =   0
         Left            =   120
         Picture         =   "frmPic.frx":1986
         Top             =   240
         Width           =   225
      End
   End
   Begin VB.Frame fraMonsters 
      Caption         =   "Monster's"
      Height          =   5055
      Left            =   3960
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      Begin VB.Image Image9 
         Height          =   255
         Left            =   120
         Top             =   2280
         Width           =   255
      End
      Begin VB.Image imgmonster 
         Height          =   840
         Index           =   9
         Left            =   1920
         Picture         =   "frmPic.frx":1D74
         Top             =   1080
         Width           =   720
      End
      Begin VB.Image imgmonster 
         Height          =   465
         Index           =   8
         Left            =   4560
         Picture         =   "frmPic.frx":23A6
         Top             =   360
         Width           =   480
      End
      Begin VB.Image imgmonster 
         Height          =   480
         Index           =   7
         Left            =   4080
         Picture         =   "frmPic.frx":285B
         Top             =   360
         Width           =   360
      End
      Begin VB.Image imgmonster 
         Height          =   480
         Index           =   6
         Left            =   3360
         Picture         =   "frmPic.frx":2CD8
         Top             =   360
         Width           =   480
      End
      Begin VB.Image imgmonster 
         Height          =   720
         Index           =   5
         Left            =   2400
         Picture         =   "frmPic.frx":31EE
         Top             =   240
         Width           =   720
      End
      Begin VB.Image imgmonster 
         Height          =   480
         Index           =   4
         Left            =   1800
         Picture         =   "frmPic.frx":379F
         Top             =   240
         Width           =   480
      End
      Begin VB.Image imgmonster 
         Height          =   960
         Index           =   3
         Left            =   720
         Picture         =   "frmPic.frx":3C5F
         Top             =   1080
         Width           =   960
      End
      Begin VB.Image imgmonster 
         Height          =   720
         Index           =   2
         Left            =   120
         Picture         =   "frmPic.frx":44C8
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image imgmonster 
         Height          =   720
         Index           =   1
         Left            =   720
         Picture         =   "frmPic.frx":4A32
         Top             =   240
         Width           =   945
      End
      Begin VB.Image imgmonster 
         Height          =   720
         Index           =   0
         Left            =   120
         Picture         =   "frmPic.frx":51C0
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Pictures"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3855
      Begin VB.Image imguser 
         Height          =   360
         Index           =   7
         Left            =   2760
         Picture         =   "frmPic.frx":572A
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imguser 
         Height          =   360
         Index           =   6
         Left            =   2400
         Picture         =   "frmPic.frx":5AFD
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imguser 
         Height          =   210
         Index           =   5
         Left            =   1920
         Picture         =   "frmPic.frx":5BDC
         Top             =   360
         Width           =   375
      End
      Begin VB.Image imguser 
         Height          =   360
         Index           =   4
         Left            =   1560
         Picture         =   "frmPic.frx":5F96
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imguser 
         Height          =   360
         Index           =   3
         Left            =   1200
         Picture         =   "frmPic.frx":6066
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imguser 
         Height          =   360
         Index           =   2
         Left            =   840
         Picture         =   "frmPic.frx":6442
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imguser 
         Height          =   360
         Index           =   1
         Left            =   480
         Picture         =   "frmPic.frx":681F
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imguser 
         Height          =   360
         Index           =   0
         Left            =   120
         Picture         =   "frmPic.frx":6BF5
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.Frame fraAllyPic 
      BackColor       =   &H8000000A&
      Caption         =   "Ally Pictures:"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      Begin VB.Image Image2 
         Height          =   360
         Index           =   9
         Left            =   3480
         Picture         =   "frmPic.frx":6FC4
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   8
         Left            =   3120
         Picture         =   "frmPic.frx":70A2
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   0
         Left            =   120
         Picture         =   "frmPic.frx":717F
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   1
         Left            =   480
         Picture         =   "frmPic.frx":725B
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   2
         Left            =   840
         Picture         =   "frmPic.frx":762A
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   210
         Index           =   3
         Left            =   1200
         Picture         =   "frmPic.frx":7A07
         Top             =   360
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   4
         Left            =   1680
         Picture         =   "frmPic.frx":7DBE
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   7
         Left            =   2760
         Picture         =   "frmPic.frx":7E9B
         Top             =   240
         Width           =   225
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   6
         Left            =   2400
         Picture         =   "frmPic.frx":7FA5
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   5
         Left            =   2040
         Picture         =   "frmPic.frx":8086
         Top             =   240
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mark - $uper-$tar
'Final Fantasy RPG
'Pics
Option Explicit
