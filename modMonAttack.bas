Attribute VB_Name = "modMonAttack"
'Mark - $uper-$tar
'Final Fantasy RPG
'Monster Attacks:
Option Explicit
'For Random Items
Dim intmonAttack As Integer 'For Random Attacks
Dim IntmonWho As Integer
Dim intmonWho1 As Integer 'For Random1 Victims
Dim intmonWho2 As Integer
Dim intmonWho3 As Integer
Dim intmonWho4 As Integer
Dim intuser As Integer 'User
Dim intally As Integer ' Ally
Dim intboth As Integer 'Both
Dim intEA As Integer 'For Monster Damage
Dim Int1 As Integer
Dim Int2 As Integer
    
'*******************[Monster Attacks]*******************'
'Monster Index:
'AC1: = 1 - Lobo
'     = 3 - Vomammoth
Public Sub Mon1Attack()
Int2 = frmMain.lblyourdead.Caption
Int1 = frmMain.lbldead.Caption
Randomize
'For Random Vicitm(User or Ally)
IntmonWho = Int1 - Int2
intmonWho1 = Int(IntmonWho * Rnd) + 1
'If 1 then = User
'If 2 then = Ally

'****Monsters****'
'Lobo:
If frmMain.lblmon1AC.Caption = 1 Then
    intmonAttack = Int((2 * Rnd) + 1)
    Select Case intmonAttack

Case 1
If intmonWho1 = 1 Then 'User:
         'Quick Attack Damage
            intEA = Int((6 * Rnd) + 1)
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack1.Caption = -intEA
         'Damage:
            frmMain.lbluserdam.Caption = intEA
         'Picture
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
        'Check
            frmMain.lblUsercheck.Caption = 0
End If

If intmonWho1 = 2 Then 'Ally
         'Quick Attack Damage
            intEA = Int((6 * Rnd) + 1)
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack1.Caption = -intEA
         'Damage:
            frmMain.lblallydam.Caption = intEA
         'Picture
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
         'Check
            frmMain.lblAllycheck.Caption = 0
    End If
Case 2
    If intmonWho1 = 1 Then 'User:
         'Tacke
            intEA = Int((5 * Rnd) + 5)
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack1.Caption = -intEA
         'Damage:
            frmMain.lbluserdam.Caption = intEA
         'Picture
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
                'Check
            frmMain.lblUsercheck.Caption = 0
    End If
    If intmonWho1 = 2 Then 'Ally
         'Takle
            intEA = Int((5 * Rnd) + 5)
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack1.Caption = -intEA
         'Damage:
            frmMain.lblallydam.Caption = intEA
         'Picture
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
        'Check
            frmMain.lblAllycheck.Caption = 0
    End If
    End Select
      End If

'Vomammoth:
If frmMain.lblmon1AC.Caption = 3 Then
    intmonAttack = Int((2 * Rnd) + 1)
    Select Case intmonAttack
    
    Case 1
        
        If intmonWho1 = 1 Then 'User:
         'Quick Attack Damage
            intEA = Int((6 * Rnd) + 6)
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack1.Caption = -intEA
         'Damage:
            frmMain.lbluserdam.Caption = intEA
         'Picture
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
        'Check
            frmMain.lblUsercheck.Caption = 0
    End If

    If intmonWho1 = 2 Then 'Ally
         'Quick Attack Damage
            intEA = Int((6 * Rnd) + 6)
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack1.Caption = -intEA
         'Damage:
            frmMain.lblallydam.Caption = intEA
         'Picture
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
         'Check
            frmMain.lblAllycheck.Caption = 0
    End If
Case 2
    If intmonWho1 = 1 Then 'User:
         'Tacke
            intEA = Int((5 * Rnd) + 15)
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack1.Caption = -intEA
         'Damage:
            frmMain.lbluserdam.Caption = intEA
         'Picture
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
                'Check
            frmMain.lblUsercheck.Caption = 0
    End If
    If intmonWho1 = 2 Then 'Ally
         'Takle
            intEA = Int((5 * Rnd) + 15)
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack1.Caption = -intEA
         'Damage:
            frmMain.lblallydam.Caption = intEA
         'Picture
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
        'Check
            frmMain.lblAllycheck.Caption = 0
    End If
End Select
    End If
End Sub
Public Sub Mon2Attack()

'Monster Index:
'AC2: = 1 - Lobo
'     = 2 - Vammoth
'
Int2 = frmMain.lblyourdead.Caption
Int1 = frmMain.lbldead.Caption
Randomize
'For Random Vicitm(User or Ally)
IntmonWho = Int1 - Int2
intmonWho2 = Int(IntmonWho * Rnd) + 1
'If 1 then = User
'If 2 then = Ally

'****Monsters****'
'Lobo:
If frmMain.lblmon2AC.Caption = 1 Then
    intmonAttack = Int((2 * Rnd) + 1)
    Select Case intmonAttack

Case 1
    If intmonWho2 = 1 Then 'User:
         'Quick Attack Damage
            intEA = Int((8 * Rnd) + 1)
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack2.Caption = -intEA
         'Damage:
            frmMain.lbluserdam.Caption = intEA
         'Picture
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
        'Check
            frmMain.lblUsercheck.Caption = 0
    End If

    If intmonWho2 = 2 Then 'Ally
         'Quick Attack Damage
            intEA = Int((8 * Rnd) + 1)
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack2.Caption = -intEA
         'Damage:
            frmMain.lblallydam.Caption = intEA
         'Picture
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
         'Check
            frmMain.lblAllycheck.Caption = 0
    End If
Case 2
    If intmonWho2 = 1 Then 'User:
         'Tacke
            intEA = Int((8 * Rnd) + 5)
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack2.Caption = -intEA
         'Damage:
            frmMain.lbluserdam.Caption = intEA
         'Picture
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
                'Check
            frmMain.lblUsercheck.Caption = 0
    End If
    If intmonWho2 = 2 Then 'Ally
         'Takle
            intEA = Int((8 * Rnd) + 5)
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack2.Caption = -intEA
         'Damage:
            frmMain.lblallydam.Caption = intEA
         'Picture
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
        'Check
            frmMain.lblAllycheck.Caption = 0
    End If
End Select
End If 'Don't Delete this! (For Attack "cases")

'Vomammoth:
If frmMain.lblmon2AC.Caption = 3 Then
    intmonAttack = Int((2 * Rnd) + 1)
    Select Case intmonAttack
    
    Case 1
        
        If intmonWho2 = 1 Then 'User:
         'Quick Attack Damage
            intEA = Int((6 * Rnd) + 6)
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack2.Caption = -intEA
         'Damage:
            frmMain.lbluserdam.Caption = intEA
         'Picture
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
        'Check
            frmMain.lblUsercheck.Caption = 0
    End If

    If intmonWho2 = 2 Then 'Ally
         'Quick Attack Damage
            intEA = Int((6 * Rnd) + 6)
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack2.Caption = -intEA
         'Damage:
            frmMain.lblallydam.Caption = intEA
         'Picture
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
         'Check
            frmMain.lblAllycheck.Caption = 0
    End If
Case 2
    If intmonWho2 = 1 Then 'User:
         'Tacke
            intEA = Int((5 * Rnd) + 15)
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack2.Caption = -intEA
         'Damage:
            frmMain.lbluserdam.Caption = intEA
         'Picture
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
                'Check
            frmMain.lblUsercheck.Caption = 0
    End If
    If intmonWho2 = 2 Then 'Ally
         'Takle
            intEA = Int((5 * Rnd) + 15)
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack2.Caption = -intEA
         'Damage:
            frmMain.lblallydam.Caption = intEA
         'Picture
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
        'Check
            frmMain.lblAllycheck.Caption = 0
    End If
End Select
    End If 'FOr AC
End Sub

Public Sub Mon3Attack()
'Monster Index:
'AC3: = 2 - Guard
'
Int2 = frmMain.lblyourdead.Caption
Int1 = frmMain.lbldead.Caption
'For Random Vicitm(User or Ally)
IntmonWho = Int1 - Int2
intmonWho3 = Int(IntmonWho * Rnd) + 1
'If 1 then = User
'If 2 then = Ally

'****Monsters****'
'Guard:
If frmMain.lblmon3AC.Caption = 2 Then
    intmonAttack = Int((2 * Rnd) + 1)
    Select Case intmonAttack

Case 1
    If intmonWho3 = 1 Then 'User:
         'Quick Attack Damage
            intEA = Int((7 * Rnd) + 5)
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack3.Caption = -intEA
         'Damage:
            frmMain.lbluserdam.Caption = intEA
         'Picture
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
        'Check
            frmMain.lblUsercheck.Caption = 0
    End If

    If intmonWho3 = 2 Then 'Ally
         'Quick Attack Damage
            intEA = Int((7 * Rnd) + 5)
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack3.Caption = -intEA
         'Damage:
            frmMain.lblallydam.Caption = intEA
         'Picture
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
         'Check
            frmMain.lblAllycheck.Caption = 0
    End If
Case 2
    If intmonWho3 = 1 Then 'User:
         'Tacke
            intEA = Int((5 * Rnd) + 10)
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack3.Caption = -intEA
         'Damage:
            frmMain.lbluserdam.Caption = intEA
         'Picture
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
                'Check
            frmMain.lblUsercheck.Caption = 0
    End If
    If intmonWho3 = 2 Then 'Ally
         'Takle
            intEA = Int((5 * Rnd) + 10)
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack3.Caption = -intEA
         'Damage:
            frmMain.lblallydam.Caption = intEA
         'Picture
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
        'Check
            frmMain.lblAllycheck.Caption = 0
    End If
End Select
End If 'Don't Delete this! (For Attack "cases")
End Sub

Public Sub Mon4Attack()
'Monster Index:
'AC3: = 2 - Guard
'
Int2 = frmMain.lblyourdead.Caption
Int1 = frmMain.lbldead.Caption
'For Random Vicitm(User or Ally)
IntmonWho = Int1 - Int2
intmonWho4 = Int(IntmonWho * Rnd) + 1
'If 1 then = User
'If 2 then = Ally

'****Monsters****'
'Guard:
If frmMain.lblmon4AC.Caption = 2 Then
    intmonAttack = Int((2 * Rnd) + 1)
    Select Case intmonAttack

Case 1
    If intmonWho4 = 1 Then 'User:
         'Quick Attack Damage
            intEA = Int((5 * Rnd) + 5)
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack4.Caption = -intEA
         'Damage:
            frmMain.lbluserdam.Caption = intEA
         'Picture
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
        'Check
            frmMain.lblUsercheck.Caption = 0
    End If
    
    If intmonWho4 = 2 Then 'Ally
         'Quick Attack Damage
            intEA = Int((5 * Rnd) + 5)
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack4.Caption = -intEA
         'Damage:
            frmMain.lblallydam.Caption = intEA
         'Picture
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
         'Check
            frmMain.lblAllycheck.Caption = 0
    End If
Case 2
    If intmonWho4 = 1 Then 'User:
         'Tacke
            intEA = Int((5 * Rnd) + 10)
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack4.Caption = -intEA
         'Damage:
            frmMain.lbluserdam.Caption = intEA
         'Picture
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
                'Check
            frmMain.lblUsercheck.Caption = 0
    End If
    If intmonWho4 = 2 Then 'Ally
         'Takle
            intEA = Int((5 * Rnd) + 10)
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intEA
         'Sets Clock and disables Attack
            frmMain.lblmonAttack4.Caption = -intEA
         'Damage:
            frmMain.lblallydam.Caption = intEA
         'Picture
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
        'Check
            frmMain.lblAllycheck.Caption = 0
    End If
End Select
End If 'Don't Delete this! (For Attack "cases")
End Sub
