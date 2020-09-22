VERSION 5.00
Begin VB.Form frmAttack 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Attacks"
   ClientHeight    =   1905
   ClientLeft      =   6225
   ClientTop       =   7275
   ClientWidth     =   2865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAttack.frx":0000
   ScaleHeight     =   1905
   ScaleWidth      =   2865
   Begin VB.ComboBox cmbitems 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   2640
      Top             =   1680
   End
   Begin VB.ComboBox cmbUMagic 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbUAttacks 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox cmbAmagic 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbAattacks 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblitem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Items:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblmagic 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Magic:"
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
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblattacks 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Attacks:"
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
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmAttack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mark - $uper-$tar
'Final Fantasy RPG
Option Explicit
Private Sub cmbAmagic_Click()
AllyMagic
End Sub

Private Sub cmbAattacks_Click()
AllyAttacks '(Uses Attacks)
End Sub

Private Sub cmbitems_Click()
PSItems
End Sub

Private Sub cmbUAttacks_Click()
UserAttacks 'Uses Attacks
End Sub

Private Sub cmbUMagic_CLick()
UserMagic 'See modspells
End Sub

Private Sub Form_Load()
    'Disables
    cmbUAttacks.Enabled = False
    cmbUMagic.Enabled = False
    cmbAattacks.Enabled = False
    cmbAmagic.Enabled = False
    cmbitems.Enabled = False
    
    'For User Attacks
        cmbUAttacks.AddItem ("Fight")
        cmbUAttacks.ListIndex = 0
    'Magic
        cmbUMagic.AddItem ("Bolt2")
        cmbUMagic.AddItem ("Fire2")
        cmbUMagic.AddItem ("Ice2")
        cmbUMagic.AddItem ("Cure")
        cmbUMagic.ListIndex = 0
    
    'For Ally Attacks
        cmbAattacks.AddItem ("Fight")
        cmbAattacks.ListIndex = 0
    'Magic
        cmbAmagic.AddItem ("Bolt")
        cmbAmagic.AddItem ("Fire")
        cmbAmagic.AddItem ("Ice")
        cmbAmagic.ListIndex = 0
        
    'Timers
        Timer1.Interval = 1
    'Text
        frmMain.lblmonhp.Caption = 27
        frmMain.lblmon1dam.Caption = ""
        frmMain.mbluserMP.Caption = "30"
        frmMain.lblstatus.Caption = ""
        frmMain.mblallymp.Caption = "15"
    
    'Items
        cmbitems.AddItem ("Potion")
        cmbitems.AddItem ("Potion")
        cmbitems.AddItem ("Potion")
        cmbitems.AddItem ("Tincture")
        cmbitems.ListIndex = 0
    
End Sub


Private Sub Timer1_Timer()
If frmMain.lblUserAttack.Caption = 60 Then
lblNAme.Caption = frmMain.lblUserName
End If
If frmMain.lblAllyAttack.Caption = 75 Then
lblNAme.Caption = frmMain.lblallyname
End If
If lblNAme.Caption = "Name:" Then
frmAttack.Visible = False
End If
End Sub

Private Sub Timer2_Timer()
If frmMain.Visible = False Then
Unload Me
End If
End Sub
