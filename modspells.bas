Attribute VB_Name = "modspells"
'Mark - $uper-$tar
'Final Fantasy RPG
'-Spells-'
Option Explicit
'Spells:
Dim intbolts As Integer ' For Bolt
Dim intfires As Integer ' For Fire
Dim intices As Integer  ' For Ice
'Magic
Public Sub AllyMagic()
        If frmMain.optnone = True Then 'Checks if there is someone to attack...
        frmMain.lblstatus.Caption = strAName & " Attack Someone!" 'Tells user if not
        Exit Sub 'Saves attack from being wasted
        End If
        
 Select Case frmAttack.cmbAmagic.Text 'Gets Spells
      Case "Bolt" '
        'Checks Magic
        If frmMain.mblallymp.Caption < 5 Then
        frmMain.lblstatus.Caption = "Not Enough Magic!"
        Exit Sub 'Saves attack from being wasted
        End If
        'Sets Bolt Damage
            If frmMain.optmon1.Value = True Then
            intbolts = intbolt1
            End If
            If frmMain.optmon2.Value = True Then
            intbolts = intbolt2
            End If
            If frmMain.optmon3.Value = True Then
            intbolts = intbolt3
            End If
            If frmMain.optmon4.Value = True Then
            intbolts = intbolt4
            End If
         'Fightin Damage:
            intUA = Int((intbolts * Rnd) + 11)
         'Resets Clock and disables Attack
            frmMain.lblAllyAttack.Caption = 0
            frmMain.shpAA.Width = 15
         'Text
            frmMain.lblstatus.Caption = strAName & " cast Bolt!"
         'Sets Damage
            If frmMain.optnone = True Then
            frmMain.lblstatus.Caption = "Attack Someone!"
            Exit Sub
            End If
            'Monster 1
            If frmMain.optmon1.Value = True Then
            frmMain.lblmonhp.Caption = frmMain.lblmonhp.Caption - intUA 'Subtract Damage
            frmMain.lblmon1check.Caption = 0 'Resets Damage Reset
            frmMain.lblmon1dam.Caption = intUA 'Shows Damage
            End If
            'Monster 2
            If frmMain.optmon2.Value = True Then
            frmMain.lblmon2dam.Caption = intUA
            frmMain.lblmon2check.Caption = 0
            frmMain.lblmon2hp.Caption = frmMain.lblmon2hp.Caption - intUA
            End If
            'Monster 3
            If frmMain.optmon3.Value = True Then
            frmMain.lblmon3dam.Caption = intUA
            frmMain.lblmon3check.Caption = 0
            frmMain.lblmon3hp.Caption = frmMain.lblmon3hp.Caption - intUA
            End If
            'Monster 4
            If frmMain.optmon4.Value = True Then
            frmMain.lblmon4check.Caption = 0
            frmMain.lblmon4dam.Caption = intUA
            frmMain.lblmon4hp.Caption = frmMain.lblmon4hp.Caption - intUA
            End If
            'Minus Magic Points:
            frmMain.mblallymp.Caption = frmMain.mblallymp.Caption - 5
            'Picture:
            frmMain.imgally.Picture = frmPic.Image2(7).Picture
            
            'Unloads Attack Screen
                frmAttack.Visible = False
                
Case "Fire"
        'Checks Magic
        If frmMain.mblallymp.Caption < 5 Then
        frmMain.lblstatus.Caption = "Not Enough Magic!"
        Exit Sub
        End If
        'Sets Bolt Damage
            If frmMain.optnone = True Then
            frmMain.lblstatus.Caption = "Attack Someone!"
            Exit Sub
            End If
            If frmMain.optmon1.Value = True Then
            intfires = intfire1
            End If
            If frmMain.optmon2.Value = True Then
            intfires = intfire2
            End If
            If frmMain.optmon3.Value = True Then
            intfires = intfire3
            End If
            If frmMain.optmon4.Value = True Then
            intfires = intfire4
            End If
         'Fightin Damage:
            intUA = Int((intfires * Rnd) + 11)
         'Resets Clock and disables Attack
            frmMain.lblAllyAttack.Caption = 0
            frmMain.shpAA.Width = 15
         'Text
            frmMain.lblstatus.Caption = strAName & " cast Fire!"
         'Sets Damage
            'Monster 1
            If frmMain.optmon1.Value = True Then
            frmMain.lblmonhp.Caption = frmMain.lblmonhp.Caption - intUA 'Subtract Damage
            frmMain.lblmon1check.Caption = 0 'Resets Damage Reset
            frmMain.lblmon1dam.Caption = intUA 'Shows Damage
            End If
            'Monster 2
            If frmMain.optmon2.Value = True Then
            frmMain.lblmon2dam.Caption = intUA
            frmMain.lblmon2check.Caption = 0
            frmMain.lblmon2hp.Caption = frmMain.lblmon2hp.Caption - intUA
            End If
            'Monster 3
            If frmMain.optmon3.Value = True Then
            frmMain.lblmon3dam.Caption = intUA
            frmMain.lblmon3check.Caption = 0
            frmMain.lblmon3hp.Caption = frmMain.lblmon3hp.Caption - intUA
            End If
            'Monster 4
            If frmMain.optmon4.Value = True Then
            frmMain.lblmon4check.Caption = 0
            frmMain.lblmon4dam.Caption = intUA
            frmMain.lblmon4hp.Caption = frmMain.lblmon4hp.Caption - intUA
            End If
            'Minus Magic Points:
            frmMain.mblallymp.Caption = frmMain.mblallymp.Caption - 5
            'Picture:
            frmMain.imgally.Picture = frmPic.Image2(7).Picture
            
            'Unloads Attack Screen
                frmAttack.Visible = False
                
Case "Ice"
        'Checks Magic
        If frmMain.mblallymp.Caption < 5 Then
        frmMain.lblstatus.Caption = "Not Enough Magic!"
        Exit Sub
        End If
        'Sets Bolt Damage
            If frmMain.optnone = True Then
            frmMain.lblstatus.Caption = "Attack Someone!"
            Exit Sub
            End If
            If frmMain.optmon1.Value = True Then
            intices = intice1
            End If
            If frmMain.optmon2.Value = True Then
            intices = intice2
            End If
            If frmMain.optmon3.Value = True Then
            intices = intice3
            End If
            If frmMain.optmon4.Value = True Then
            intices = intice4
            End If
         'Fightin Damage:
            intUA = Int((intices * Rnd) + 11)
         'Resets Clock and disables Attack
            frmMain.lblAllyAttack.Caption = 0
            frmMain.shpAA.Width = 15
         'Text
            frmMain.lblstatus.Caption = strAName & " cast Ice!"
         'Sets Damage
            'Monster 1
            If frmMain.optmon1.Value = True Then
            frmMain.lblmonhp.Caption = frmMain.lblmonhp.Caption - intUA 'Subtract Damage
            frmMain.lblmon1check.Caption = 0 'Resets Damage Reset
            frmMain.lblmon1dam.Caption = intUA 'Shows Damage
            End If
            'Monster 2
            If frmMain.optmon2.Value = True Then
            frmMain.lblmon2dam.Caption = intUA
            frmMain.lblmon2check.Caption = 0
            frmMain.lblmon2hp.Caption = frmMain.lblmon2hp.Caption - intUA
            End If
            'Monster 3
            If frmMain.optmon3.Value = True Then
            frmMain.lblmon3dam.Caption = intUA
            frmMain.lblmon3check.Caption = 0
            frmMain.lblmon3hp.Caption = frmMain.lblmon3hp.Caption - intUA
            End If
            'Monster 4
            If frmMain.optmon4.Value = True Then
            frmMain.lblmon4check.Caption = 0
            frmMain.lblmon4dam.Caption = intUA
            frmMain.lblmon4hp.Caption = frmMain.lblmon4hp.Caption - intUA
            End If
            'Minus Magic Points:
            frmMain.mblallymp.Caption = frmMain.mblallymp.Caption - 5
            'Picture:
            frmMain.imgally.Picture = frmPic.Image2(7).Picture
            
            'Unloads Attack Screen
                frmAttack.Visible = False
        End Select
End Sub



Public Sub UserMagic()
            If frmMain.optnone = True Then
            frmMain.lblstatus.Caption = strAName & " Attack Someone!"
            Exit Sub
            End If
    Select Case frmAttack.cmbUMagic.Text
      Case "Bolt2"
        'Checks Magic
        If frmMain.mbluserMP.Caption < 7 Then
        frmMain.lblstatus.Caption = "Not Enough Magic!"
        Exit Sub
        End If
        'Sets Bolt Damage
            If frmMain.optnone = True Then
            frmMain.lblstatus.Caption = "Attack Someone!"
            Exit Sub
            End If
            If frmMain.optmon1.Value = True Then
            intbolts = intbolt1
            End If
            If frmMain.optmon2.Value = True Then
            intbolts = intbolt2
            End If
            If frmMain.optmon3.Value = True Then
            intbolts = intbolt3
            End If
            If frmMain.optmon4.Value = True Then
            intbolts = intbolt4
            End If
         'Fightin Damage:
            intUA = Int((intbolts * Rnd) + 41)
         'Resets Clock and disables Attack
            frmMain.lblUserAttack.Caption = 0
            frmMain.shpUA.Width = 15
         'Text
            frmMain.lblstatus.Caption = strName & " cast Bolt2!"
         'Sets Damage
            'Monster 1
            If frmMain.optmon1.Value = True Then
            frmMain.lblmonhp.Caption = frmMain.lblmonhp.Caption - intUA 'Subtract Damage
            frmMain.lblmon1check.Caption = 0 'Resets Damage Reset
            frmMain.lblmon1dam.Caption = intUA 'Shows Damage
            End If
            'Monster 2
            If frmMain.optmon2.Value = True Then
            frmMain.lblmon2dam.Caption = intUA
            frmMain.lblmon2check.Caption = 0
            frmMain.lblmon2hp.Caption = frmMain.lblmon2hp.Caption - intUA
            End If
            'Monster 3
            If frmMain.optmon3.Value = True Then
            frmMain.lblmon3dam.Caption = intUA
            frmMain.lblmon3check.Caption = 0
            frmMain.lblmon3hp.Caption = frmMain.lblmon3hp.Caption - intUA
            End If
            'Monster 4
            If frmMain.optmon4.Value = True Then
            frmMain.lblmon4check.Caption = 0
            frmMain.lblmon4dam.Caption = intUA
            frmMain.lblmon4hp.Caption = frmMain.lblmon4hp.Caption - intUA
            End If
            'Minus Magic Points:
            frmMain.mbluserMP.Caption = frmMain.mbluserMP.Caption - 7
            'Picture:
            frmMain.imguser.Picture = frmPic.imguser(7).Picture
            
            'Unloads Attack Screen
                frmAttack.Visible = False
'Fire2:
        Case "Fire2"
        'Checks Magic
        If frmMain.mbluserMP.Caption < 7 Then
        frmMain.lblstatus.Caption = "Not Enough Magic!"
        Exit Sub
        End If
        'Sets Fire Damage
            If frmMain.optnone = True Then
            frmMain.lblstatus.Caption = "Attack Someone!"
            Exit Sub
            End If
            If frmMain.optmon1.Value = True Then
            intfires = intfire1
            End If
            If frmMain.optmon2.Value = True Then
            intfires = intfire2
            End If
            If frmMain.optmon3.Value = True Then
            intfires = intfire3
            End If
            If frmMain.optmon4.Value = True Then
            intfires = intfire4
            End If
         'Fightin Damage:
            intUA = Int((intfires * Rnd) + 41)
         'Resets Clock and disables Attack
            frmMain.lblUserAttack.Caption = 0
            frmMain.shpUA.Width = 15
         'Text
            frmMain.lblstatus.Caption = strName & " cast Fire2!"
         'Sets Damage
            'Monster 1
            If frmMain.optmon1.Value = True Then
            frmMain.lblmonhp.Caption = frmMain.lblmonhp.Caption - intUA 'Subtract Damage
            frmMain.lblmon1check.Caption = 0 'Resets Damage Reset
            frmMain.lblmon1dam.Caption = intUA 'Shows Damage
            End If
            'Monster 2
            If frmMain.optmon2.Value = True Then
            frmMain.lblmon2dam.Caption = intUA
            frmMain.lblmon2check.Caption = 0
            frmMain.lblmon2hp.Caption = frmMain.lblmon2hp.Caption - intUA
            End If
            'Monster 3
            If frmMain.optmon3.Value = True Then
            frmMain.lblmon3dam.Caption = intUA
            frmMain.lblmon3check.Caption = 0
            frmMain.lblmon3hp.Caption = frmMain.lblmon3hp.Caption - intUA
            End If
            'Monster 4
            If frmMain.optmon4.Value = True Then
            frmMain.lblmon4check.Caption = 0
            frmMain.lblmon4dam.Caption = intUA
            frmMain.lblmon4hp.Caption = frmMain.lblmon4hp.Caption - intUA
            End If
            'Minus Magic Points:
            frmMain.mbluserMP.Caption = frmMain.mbluserMP.Caption - 7
            'Picture:
            frmMain.imguser.Picture = frmPic.imguser(7).Picture
            
            'Unloads Attack Screen
                frmAttack.Visible = False
                
        'Ice2:
        Case "Ice2"
        'Checks Magic
            If frmMain.optnone = True Then
            frmMain.lblstatus.Caption = "Attack Someone!"
            Exit Sub
            End If
        If frmMain.mbluserMP.Caption < 7 Then
        frmMain.lblstatus.Caption = "Not Enough Magic!"
        Exit Sub
        End If
        'Sets Fire Damage
            If frmMain.optmon1.Value = True Then
            intices = intice1
            End If
            If frmMain.optmon2.Value = True Then
            intices = intice2
            End If
            If frmMain.optmon3.Value = True Then
            intices = intice3
            End If
            If frmMain.optmon4.Value = True Then
            intices = intice4
            End If
         'Fightin Damage:
            intUA = Int((intfires * Rnd) + 41)
         'Resets Clock and disables Attack
            frmMain.lblUserAttack.Caption = 0
            frmMain.shpUA.Width = 15
         'Text
            frmMain.lblstatus.Caption = strName & " cast Ice2!"
         'Sets Damage
            'Monster 1
            If frmMain.optmon1.Value = True Then
            frmMain.lblmonhp.Caption = frmMain.lblmonhp.Caption - intUA 'Subtract Damage
            frmMain.lblmon1check.Caption = 0 'Resets Damage Reset
            frmMain.lblmon1dam.Caption = intUA 'Shows Damage
            End If
            'Monster 2
            If frmMain.optmon2.Value = True Then
            frmMain.lblmon2dam.Caption = intUA
            frmMain.lblmon2check.Caption = 0
            frmMain.lblmon2hp.Caption = frmMain.lblmon2hp.Caption - intUA
            End If
            'Monster 3
            If frmMain.optmon3.Value = True Then
            frmMain.lblmon3dam.Caption = intUA
            frmMain.lblmon3check.Caption = 0
            frmMain.lblmon3hp.Caption = frmMain.lblmon3hp.Caption - intUA
            End If
            'Monster 4
            If frmMain.optmon4.Value = True Then
            frmMain.lblmon4check.Caption = 0
            frmMain.lblmon4dam.Caption = intUA
            frmMain.lblmon4hp.Caption = frmMain.lblmon4hp.Caption - intUA
            End If
            'Minus Magic Points:
            frmMain.mbluserMP.Caption = frmMain.mbluserMP.Caption - 7
            'Picture:
            frmMain.imguser.Picture = frmPic.imguser(7).Picture
            
            'Unloads Attack Screen
                frmAttack.Visible = False
        End Select
End Sub
