Attribute VB_Name = "modattacks"
'Mark - $uper-$tar
'Final Fantasy RPG
'User/Attacks
Option Explicit
Public intUA As Integer
Public strName As String
Public strAName As String

'Ally Attacks:
Public Sub AllyAttacks()
            If frmMain.optnone = True Then
            frmMain.lblstatus.Caption = strAName & " Attack Someone!"
            Exit Sub
            End If

strAName = frmMain.lblallyname.Caption

    Select Case frmAttack.cmbAattacks.Text
      Case "Fight"
    'Moves Guy
    frmMain.imgally.Left = 5400
            'Fightin Damage:
            intUA = Int((5 * Rnd) + 35)
         'Resets Clock and disables Attack
            frmMain.lblAllyAttack.Caption = 0
            frmMain.shpAA.Width = 15
         'Text
            frmMain.lblstatus.Caption = strAName & " attacks!"
        'Sets Damage
            If frmMain.optzuser.Value = True Then
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intUA 'Subtract Damage
            frmMain.lbluserdam.Caption = intUA 'Shows Damage
            frmMain.imguser.Picture = frmPic.imguser(3).Picture
            frmMain.lblUsercheck.Caption = 0
            End If
            If frmMain.optzally.Value = True Then
            frmMain.lblstatus.Caption = strAName & " hit himself!"
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intUA 'Subtract Damage
            frmMain.lblallydam.Caption = intUA 'Shows Damage
            frmMain.imgally.Picture = frmPic.Image2(6).Picture
            frmMain.lblAllycheck.Caption = 0
            End If
            If frmMain.optmon1.Value = True Then
            frmMain.lblmonhp.Caption = frmMain.lblmonhp.Caption - intUA 'Subtract Damage
            frmMain.lblmon1check.Caption = 0 'Resets Damage Reset
            frmMain.lblmon1dam.Caption = intUA 'Shows Damage
            End If
            If frmMain.optmon2.Value = True Then
            frmMain.lblmon2dam.Caption = intUA
            frmMain.lblmon2check.Caption = 0
            frmMain.lblmon2hp.Caption = frmMain.lblmon2hp.Caption - intUA
            End If
            If frmMain.optmon3.Value = True Then
            frmMain.lblmon3dam.Caption = intUA
            frmMain.lblmon3check.Caption = 0
            frmMain.lblmon3hp.Caption = frmMain.lblmon3hp.Caption - intUA
            End If
            If frmMain.optmon4.Value = True Then
            frmMain.lblmon4check.Caption = 0
            frmMain.lblmon4dam.Caption = intUA
            frmMain.lblmon4hp.Caption = frmMain.lblmon4hp.Caption - intUA
            End If
            
            'Unloads Attack Screen
                        frmAttack.Visible = False
        
   End Select

End Sub

'User Attacks:
Public Sub UserAttacks()
            If frmMain.optnone = True Then
            frmMain.lblstatus.Caption = strAName & " Attack Someone!"
            Exit Sub
            End If
strName = frmMain.lblUserName.Caption

    Select Case frmAttack.cmbUAttacks.Text
      Case "Fight"
          'Moves Guy
            frmMain.imguser.Left = 5280
         'Fightin Damage:
            intUA = Int((10 * Rnd) + 55)
         'Resets Clock and disables Attack
            frmMain.lblUserAttack.Caption = 0
            frmMain.shpUA.Width = 15
         'Text
            frmMain.lblstatus.Caption = strName & " attacks!"
        'Sets Damage
            If frmMain.optzuser.Value = True Then
            frmMain.lblstatus.Caption = strName & " hits himself"
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption - intUA 'Subtract Damage
            frmMain.lbluserdam.Caption = intUA 'Shows Damage
            frmMain.imguser.Picture = frmPic.imguser(6).Picture
            frmMain.lblUsercheck.Caption = 0
            End If
            If frmMain.optzally.Value = True Then
            frmMain.lblallylife.Caption = frmMain.lblallylife.Caption - intUA 'Subtract Damage
            frmMain.lblallydam.Caption = intUA 'Shows Damage
            frmMain.imgally.Picture = frmPic.Image2(4).Picture
            frmMain.lblAllycheck.Caption = 0
            End If
            If frmMain.optmon1.Value = True Then
            frmMain.lblmonhp.Caption = frmMain.lblmonhp.Caption - intUA 'Subtract Damage
            frmMain.lblmon1check.Caption = 0 'Resets Damage Reset
            frmMain.lblmon1dam.Caption = intUA 'Shows Damage
            End If
            If frmMain.optmon2.Value = True Then
            frmMain.lblmon2dam.Caption = intUA
            frmMain.lblmon2check.Caption = 0
            frmMain.lblmon2hp.Caption = frmMain.lblmon2hp.Caption - intUA
            End If
            If frmMain.optmon3.Value = True Then
            frmMain.lblmon3dam.Caption = intUA
            frmMain.lblmon3check.Caption = 0
            frmMain.lblmon3hp.Caption = frmMain.lblmon3hp.Caption - intUA
            End If
            If frmMain.optmon4.Value = True Then
            frmMain.lblmon4check.Caption = 0
            frmMain.lblmon4dam.Caption = intUA
            frmMain.lblmon4hp.Caption = frmMain.lblmon4hp.Caption - intUA
            End If
            'Unloads Attack Screen
            frmAttack.Visible = False
        
   End Select

End Sub

