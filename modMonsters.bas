Attribute VB_Name = "modMonsters"
'Mark - $uper-$tar
'Final Fantasy RPG
'Monsters:
Option Explicit

Dim intradmon As Integer 'For Random Monster
'For Monster 1
Public strmon1Name As String
Public intmon1hp As Integer
Public intmon1mp As Integer
'For Spells
Public intbolt1 As Integer
Public intfire1 As Integer
Public intice1 As Integer
'For Monster 2
Public strmon2Name As String
Public intmon2hp As Integer
Public intmon2mp As Integer
'For Spells
Public intbolt2 As Integer
Public intfire2 As Integer
Public intice2 As Integer
'For Monster 3
Public strmon3name As String
Public intmon3hp As Integer
Public intmon3mp As Integer
'For Spells
Public intbolt3 As Integer
Public intfire3 As Integer
Public intice3 As Integer
'For Monster 4
Public strmon4Name As String
Public intmon4hp As Integer
Public intmon4mp As Integer
'For Spells
Public intbolt4 As Integer
Public intfire4 As Integer
Public intice4 As Integer


Public Sub randommonster() 'Random Monsters
    'Reset Timers
    frmMain.timMon1.Enabled = True
    frmMain.timMon2.Enabled = True
    frmMain.timMon3.Enabled = True
    frmMain.timMon4.Enabled = True
    frmMain.timdamageset.Enabled = True
    'Reset Values
    frmMain.lbldmagecount.Caption = 0
    frmMain.lblmonAttack1.Caption = 0
    frmMain.lblmonAttack2.Caption = 0
    frmMain.lblmonAttack3.Caption = 0
    frmMain.lblmonAttack4.Caption = 0
Randomize
intradmon = Int((2 * Rnd) + 1)
Select Case intradmon
        
    Case 1 '2 Guards, 2 Wolves
    'Monster 1:
    'Name of Monster: Lobo
        strmon1Name = "Lobo"
        frmMain.lblmonName.Caption = "Name:  " & strmon1Name
    'Hp(Life/Max)
        intmon1hp = Int((10 * Rnd) + 25)
        frmMain.lblmonhp.Caption = intmon1hp
    'Mp(Magic/Max)
        intmon1mp = 0
        frmMain.lblmonmp.Caption = intmon1mp
    'Monster Attack Case:
        frmMain.lblmon1AC.Caption = 1
    'Picture
        frmMain.imgmon1.Picture = frmPic.imgmonster(5).Picture
    'Set Case:
        frmMain.lblmonMC.Caption = 1
    'Magic Damage:
        intbolt1 = "50"
        intfire1 = "50"
        intice1 = "50"
    'Text
        frmMain.lblmon1N.Caption = "Lobo"
        
    'Monster 2
    'Name of Monster: Lobo
        strmon2Name = "Lobo"
        frmMain.lblmon2name.Caption = "Name:  " & strmon2Name
    'Hp(Life/Max)
        intmon2hp = Int((10 * Rnd) + 25)
        frmMain.lblmon2hp.Caption = intmon2hp
    'Mp(Magic/Max)
        intmon2mp = 0
        frmMain.lblmon2mp.Caption = intmon2mp
    'Monster Attack Case:
        frmMain.lblmon2AC.Caption = 1
    'Picture
        frmMain.imgmon2.Picture = frmPic.imgmonster(5).Picture
    'Set Case:
        frmMain.lblmonMC2.Caption = 1
    'Magic Damage:
        intbolt2 = "50"
        intfire2 = "50"
        intice2 = "50"
    'Text
        frmMain.lblmon2N.Caption = "Lobo"
    
    'Monster 3:
    'Name of Monster: Guard
        strmon3name = "Guard"
        frmMain.lblmon3name.Caption = "Name:  " & strmon3name
    'Hp(Life/Max)
        intmon3hp = Int((20 * Rnd) + 35)
        frmMain.lblmon3hp.Caption = intmon3hp
    'Mp(Magic/Max)
        intmon3mp = 0
        frmMain.lblmon3mp.Caption = intmon3mp
    'Monster Attack Case:
        frmMain.lblmon3AC.Caption = 2
    'Picture
        frmMain.imgmon3.Picture = frmPic.imgmonster(0).Picture
    'Set Case:
        frmMain.lblmonMC3.Caption = 2
    'Magic Damage:
        intbolt3 = "50"
        intfire3 = "50"
        intice3 = "50"
    'Text
        frmMain.lblmon3N.Caption = "Guard"
        
    'Monster 4:
    'Name of Monster: Guard
        strmon4Name = "Guard"
        frmMain.lblmon4name.Caption = "Name:  " & strmon4Name
    'Hp(Life/Max)
        intmon4hp = Int((20 * Rnd) + 35)
        frmMain.lblmon4hp.Caption = intmon4hp
    'Mp(Magic/Max)
        intmon4mp = 0
        frmMain.lblmon4mp.Caption = intmon4mp
    'Monster Attack Case:
        frmMain.lblmon4AC.Caption = 2
    'Picture
        frmMain.imgmon4.Picture = frmPic.imgmonster(0).Picture
    'Set Case:
        frmMain.lblmonMC4.Caption = 2
    'Magic Damage:
        intbolt4 = "50"
        intfire4 = "50"
        intice4 = "50"
    'Text
        frmMain.lblmon4N.Caption = "Guard"
        
    'All (Dead):
        frmMain.lblalldead.Caption = 4
    
    'All (exp/gold)
    frmMain.lblexp.Caption = Int(56 * Rnd) + 146
    frmMain.lblgold.Caption = Int(300 * Rnd) + 104
    
    Case 2 '2 Guards, 2 Vomammoths
    'Monster 1:
    'Name of Monster: Vomammoth
        strmon1Name = "Vomammoth"
        frmMain.lblmonName.Caption = "Name:  " & strmon1Name
    'Hp(Life/Max)
        intmon1hp = Int((20 * Rnd) + 90)
        frmMain.lblmonhp.Caption = intmon1hp
    'Mp(Magic/Max)
        intmon1mp = 30
        frmMain.lblmonmp.Caption = intmon1mp
    'Monster Attack Case:
        frmMain.lblmon1AC.Caption = 3
    'Picture
        frmMain.imgmon1.Picture = frmPic.imgmonster(1).Picture
    'Set Case:
        frmMain.lblmonMC.Caption = 3
    'Magic Damage:
        intbolt1 = "50"
        intfire1 = "75"
        intice1 = "25"
    'Text
        frmMain.lblmon1N.Caption = "Vomammoth"
        
    'Monster 2
    'Name of Monster: Lobo
        strmon2Name = "Vomammoth"
        frmMain.lblmon2name.Caption = "Name:  " & strmon2Name
    'Hp(Life/Max)
        intmon2hp = Int((20 * Rnd) + 90)
        frmMain.lblmon2hp.Caption = intmon2hp
    'Mp(Magic/Max)
        intmon2mp = 30
        frmMain.lblmon2mp.Caption = intmon2mp
    'Monster Attack Case:
        frmMain.lblmon2AC.Caption = 3
    'Picture
        frmMain.imgmon2.Picture = frmPic.imgmonster(1).Picture
    'Set Case:
        frmMain.lblmonMC2.Caption = 1
    'Magic Damage:
        intbolt2 = "50"
        intfire2 = "50"
        intice2 = "50"
    'Text
        frmMain.lblmon2N.Caption = "Vomammoth"
    
    'Monster 3:
    'Name of Monster: Guard
        strmon3name = "Guard"
        frmMain.lblmon3name.Caption = "Name:  " & strmon3name
    'Hp(Life/Max)
        intmon3hp = Int((20 * Rnd) + 35)
        frmMain.lblmon3hp.Caption = intmon3hp
    'Mp(Magic/Max)
        intmon3mp = 0
        frmMain.lblmon3mp.Caption = intmon3mp
    'Monster Attack Case:
        frmMain.lblmon3AC.Caption = 2
    'Picture
        frmMain.imgmon3.Picture = frmPic.imgmonster(0).Picture
    'Set Case:
        frmMain.lblmonMC3.Caption = 2
    'Magic Damage:
        intbolt3 = "50"
        intfire3 = "50"
        intice3 = "50"
    'Text
        frmMain.lblmon3N.Caption = "Guard"
        
    'Monster 4:
    'Name of Monster: Guard
        strmon4Name = "Guard"
        frmMain.lblmon4name.Caption = "Name:  " & strmon4Name
    'Hp(Life/Max)
        intmon4hp = Int((20 * Rnd) + 35)
        frmMain.lblmon4hp.Caption = intmon4hp
    'Mp(Magic/Max)
        intmon4mp = 0
        frmMain.lblmon4mp.Caption = intmon4mp
    'Monster Attack Case:
        frmMain.lblmon4AC.Caption = 2
    'Picture
        frmMain.imgmon4.Picture = frmPic.imgmonster(0).Picture
    'Set Case:
        frmMain.lblmonMC4.Caption = 2
    'Magic Damage:
        intbolt4 = "50"
        intfire4 = "50"
        intice4 = "50"
    'Text
        frmMain.lblmon4N.Caption = "Guard"
        
    'All (Dead):
        frmMain.lblalldead.Caption = 4
    
    'All (exp/gold)
    frmMain.lblexp.Caption = Int(56 * Rnd) + 146
    frmMain.lblgold.Caption = Int(300 * Rnd) + 104
    End Select
    End Sub

