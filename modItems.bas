Attribute VB_Name = "modItems"
'Mark - $uper-$tar
'Final Fantasy RPG
'Monsters:
Option Explicit
'For Random Items
Dim intraditem As Integer 'For Random Item
    
'*******************[Items]*******************'
 Public Sub PSItems()
    Select Case frmAttack.cmbitems.Text
    
    Case "Potion"
        If frmMain.optzally.Value = True Then
            If frmMain.lblallylife.Caption = "0" Then
            frmMain.lblallylife.Caption = "0"
            Exit Sub
        End If
        'Ally Life:
                frmMain.lblallylife.Caption = frmMain.lblallylife.Caption + 100
            If frmMain.lblallylife.Caption > frmMain.lblAllyhp.Caption Then
                frmMain.lblallylife.Caption = frmMain.lblAllyhp.Caption
            End If
         'Resets Clock and disables Attack
            frmMain.lblAllyAttack.Caption = 0
            frmMain.shpAA.Width = 15
         'Text
            frmMain.lblstatus.Caption = "Potion"
         'Remove Potion
                frmAttack.cmbitems.RemoveItem frmAttack.cmbitems.ListIndex
         'Unloads Attack Screen
            frmAttack.Visible = False
                'Picture
            frmMain.imgally.Picture = frmPic.Image2(1).Picture
        End If
        
        'User Life:
            If frmMain.optzuser.Value = True Then
        'Add Life
            frmMain.lbluserlife.Caption = frmMain.lbluserlife.Caption + 100
        If frmMain.lbluserlife.Caption > frmMain.lblUMaxLife.Caption Then
            frmMain.lbluserlife.Caption = frmMain.lblUMaxLife.Caption
        End If
         'Resets Clock and disables Attack
            frmMain.lblUserAttack.Caption = 0
            frmMain.shpUA.Width = 15
         'Text
            frmMain.lblstatus.Caption = "Potion"
         'Remove Potion
                frmAttack.cmbitems.RemoveItem frmAttack.cmbitems.ListIndex
         'Unloads Attack Screen
            frmAttack.Visible = False
        'Picture
            frmMain.imguser.Picture = frmPic.imguser(0).Picture
        End If
        
        Case "Tincture"
        'Ally Magic:
        
                frmMain.mblallymp.Caption = frmMain.mblallymp.Caption + 100
            If frmMain.mblallymp.Caption > frmMain.mblallympMax.Caption Then
                frmMain.mblallymp.Caption = frmMain.mblallympMax.Caption
            End If
         'Resets Clock and disables Attack
            frmMain.lblAllyAttack.Caption = 0
            frmMain.shpAA.Width = 15
         'Text
            frmMain.lblstatus.Caption = "Tincture"
         'Remove "Tincture"
                frmAttack.cmbitems.RemoveItem frmAttack.cmbitems.ListIndex
         'Unloads Attack Screen
            frmAttack.Visible = False
        
        
        'User "Tincture":
            
            If frmMain.optzuser.Value = True Then
        'Add Magic
            frmMain.mbluserMP.Caption = frmMain.mbluserMP.Caption + 50
        If frmMain.mbluserMP.Caption > frmMain.mbluserMaxmp.Caption Then
            frmMain.mbluserMP.Caption = frmMain.mbluserMaxmp.Caption
        End If
         'Resets Clock and disables Attack
            frmMain.lblUserAttack.Caption = 0
            frmMain.shpUA.Width = 15
         'Text
            frmMain.lblstatus.Caption = "Tincture"
         'Remove "Tincture"
                frmAttack.cmbitems.RemoveItem frmAttack.cmbitems.ListIndex
         'Unloads Attack Screen
            frmAttack.Visible = False
        End If
        
        Case "Phoniex Down"
            If frmMain.lblallylife.Caption = "0" Then
            'Cure
            frmMain.lblallylife.Caption = "25"
            'Text
             frmMain.lblstatus.Caption = "Phoniex Down"
            'Able to fight
            frmMain.timallyattacks.Enabled = True
            'Picture
            frmMain.imgally.Picture = frmPic.Image2(5).Picture
            'Remove Potion
            frmAttack.cmbitems.RemoveItem frmAttack.cmbitems.ListIndex
            'Fix to life:
            frmMain.lbldead.Caption = 3
            'Resets Clock and disables Attack
            frmMain.lblUserAttack.Caption = 0
            frmMain.shpUA.Width = 15
            Else
             Exit Sub
            End If
        End Select
End Sub

Public Sub randomitems() 'When user wins, they get 1 item:
Randomize
intraditem = Int((3 * Rnd) + 1)
Select Case intraditem
    
    Case 1 'Potion
    frmAttack.cmbitems.AddItem ("Potion")
     frmBattleWin.lblitem.Caption = "Recived a Potion"

    Case 2 'Tincture
    frmAttack.cmbitems.AddItem ("Tincture")
    frmBattleWin.lblitem.Caption = "Recived a Tincture"
    
    
    Case 3 'Phoniex
    frmAttack.cmbitems.AddItem ("Phoniex Down")
    frmBattleWin.lblitem.Caption = "Recived a Phoniex Down"

End Select
End Sub



