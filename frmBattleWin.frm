VERSION 5.00
Begin VB.Form frmBattleWin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "You Won!"
   ClientHeight    =   1875
   ClientLeft      =   7140
   ClientTop       =   5115
   ClientWidth     =   2385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmBattleWin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBattleWin.frx":08CA
   ScaleHeight     =   1875
   ScaleWidth      =   2385
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblitem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblgold 
      BackStyle       =   0  'Transparent
      Caption         =   "Gold:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblexp 
      BackStyle       =   0  'Transparent
      Caption         =   "EXP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Final Fantay RPG"
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
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmBattleWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mark - $uper-$tar
'Final Fantasy RPG
'Battle Win
Option Explicit
Dim intexp As Integer
Dim intgold As Integer

Private Sub Form_Load()
'Items:
    randomitems

'Status:
    frmMain.lblstatus.Caption = "You've Won!"

'For Exp:
    intexp = frmMain.lblexp.Caption
    lblexp.Caption = "Exp: " & intexp
'For Gold:
    intgold = frmMain.lblgold.Caption
    lblgold.Caption = "Gold: " & intgold
End Sub

Public Sub ResetBattle()

    'reset- optmon's attack
    frmMain.optmon1.Visible = True
    frmMain.optmon2.Visible = True
    frmMain.optmon3.Visible = True
    frmMain.optmon4.Visible = True
    'Reset Pictures
    frmMain.imgmon1.Visible = True
    frmMain.imgmon2.Visible = True
    frmMain.imgmon3.Visible = True
    frmMain.imgmon4.Visible = True
    'Reset Text
    frmMain.lblmon1N.Visible = True
    frmMain.lblmon2N.Visible = True
    frmMain.lblmon3N.Visible = True
    frmMain.lblmon4N.Visible = True
    'monster
    randommonster
        'to play again:
        frmMain.timwinlose.Enabled = True
        frmMain.optmon1.SetFocus
        'Resets
        frmMain.lblmon1dam.Visible = True
        frmMain.lblmon2dam.Visible = True
        frmMain.lblmon3dam.Visible = True
        frmMain.lblmon4dam.Visible = True
        'Resets
        Unload Me
End Sub

Private Sub Form_Click()
ResetBattle
End Sub


Private Sub lblexp_Click()
ResetBattle
End Sub

Private Sub lblgold_Click()
ResetBattle
End Sub

Private Sub lblitem_Click()
ResetBattle
End Sub

Private Sub lblName_Click()
ResetBattle
End Sub
