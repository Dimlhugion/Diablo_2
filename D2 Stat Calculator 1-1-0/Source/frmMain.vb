'Dimlhugion
'D2 Stat Calculator
'Last Revised: 6/16/16
'Current Release: 1.1.0
'Current Beta: 1.1.0
'Working on: 

Public Class frmMain
    Dim intStartStrength, intStartDex, intStartVit, intStartEnergy, intStartLife, intStartStamina, intStartMana As Integer
    Dim intAddedStrength, intAddedDex, intAddedVit, intAddedEnergy, intAddedLife, intAddedSkill, intAddedStat As Integer
    Dim intCharLevel, intEndStrength, intEndDex, intEndVit, intEndEnergy, intEndSkill, intEndStat, intAvailableStat As Integer
    Dim decPerLevelLife, decPerLevelStam, decPerLevelMana, decEndLife, decEndStamina, decEndMana As Decimal
    Dim decPerVitLife, decPerVitStam, decPerEnergyMana As Decimal


    Private Sub Refresh_Data()
        Dim intLevelSkill, intLevelStat As Integer
        Dim decLevelLife, decLevelStam, decLevelMana, decPointLife, decPointStam, decPointMana As Decimal

        'Calculate level gains
        decLevelLife = decPerLevelLife * intCharLevel
        decLevelStam = decPerLevelStam * intCharLevel
        decLevelMana = decPerLevelMana * intCharLevel
        intLevelSkill = intCharLevel * 1
        intLevelStat = intCharLevel * 5

        'Calculate life/stam/mana per stat point gains
        decPointLife = decPerVitLife * intAddedVit
        decPointStam = decPerVitStam * intAddedVit
        decPointMana = decPerEnergyMana * intAddedEnergy

        'Calculate Final attributes
        intEndStrength = intStartStrength + intAddedStrength
        intEndDex = intStartDex + intAddedDex
        intEndVit = intStartVit + intAddedVit
        intEndEnergy = intStartEnergy + intAddedEnergy
        decEndLife = intStartLife + decLevelLife + intAddedLife + decPointLife
        decEndStamina = intStartStamina + decLevelStam + decPointStam
        decEndMana = intStartMana + decLevelMana + decPointMana
        intEndSkill = intLevelSkill + intAddedSkill
        intEndStat = intLevelStat + intAddedStat

        'Calculate Available Stat Points
        intAvailableStat = intEndStat - (intAddedStrength + intAddedDex + intAddedVit + intAddedEnergy)

        'Show starting attributes
        lblStartDex.Text = intStartDex.ToString("N0")
        lblStartEnergy.Text = intStartEnergy.ToString("N0")
        lblStartLife.Text = intStartLife.ToString("N0")
        lblStartMana.Text = intStartMana.ToString("N0")
        lblStartStam.Text = intStartStamina.ToString("N0")
        lblStartStrength.Text = intStartStrength.ToString("N0")
        lblStartVit.Text = intStartVit.ToString("N0")

        'Show level gains
        lblLevelLife.Text = decLevelLife.ToString("N")
        lblLevelStam.Text = decLevelStam.ToString("N")
        lblLevelMana.Text = decLevelMana.ToString("N")
        lblLevelSkill.Text = intLevelSkill.ToString("N0")
        lblLevelStat.Text = intLevelStat.ToString("N0")

        'Show points Added
        lblAddStrength.Text = intAddedStrength.ToString("N0")
        lblAddDex.Text = intAddedDex.ToString("N0")
        lblAddVit.Text = intAddedVit.ToString("N0")
        lblAddEnergy.Text = intAddedEnergy.ToString("N0")
        lblAddLife.Text = intAddedLife.ToString("N0")
        lblAddSkill.Text = intAddedSkill.ToString("N0")
        lblAddStat.Text = intAddedStat.ToString("N0")

        'Show final attributes
        lblEndStrength.Text = intEndStrength.ToString("N0")
        lblEndDex.Text = intEndDex.ToString("N0")
        lblEndVit.Text = intEndVit.ToString("N0")
        lblEndEnergy.Text = intEndEnergy.ToString("N0")
        lblEndLife.Text = decEndLife.ToString("N")
        lblEndStam.Text = decEndStamina.ToString("N")
        lblEndMana.Text = decEndMana.ToString("N")
        lblEndSkill.Text = intEndSkill.ToString("N0")
        lblEndStat.Text = intEndStat.ToString("N0")

        'Show available stat points
        lblAvailableStat.Text = intAvailableStat.ToString("N0")
    End Sub

    Private Sub Class_Select_Event(sender As Object, e As EventArgs) Handles cmbClasses.SelectedIndexChanged
        'Assign starting attributes based on class selected
        Select Case cmbClasses.SelectedItem
            Case "Amazon"
                intStartStrength = 20
                intStartDex = 25
                intStartVit = 20
                intStartEnergy = 15
                intStartLife = 50
                intStartStamina = 84
                intStartMana = 15
                decPerLevelLife = 2
                decPerLevelStam = 1
                decPerLevelMana = 1.5
                decPerVitLife = 3
                decPerVitStam = 1
                decPerEnergyMana = 1.5
            Case "Assassin"
                intStartStrength = 20
                intStartDex = 20
                intStartVit = 20
                intStartEnergy = 25
                intStartLife = 50
                intStartStamina = 95
                intStartMana = 25
                decPerLevelLife = 2
                decPerLevelStam = 1.25
                decPerLevelMana = 1.5
                decPerVitLife = 3
                decPerVitStam = 1.25
                decPerEnergyMana = 1.75
            Case "Barbarian"
                intStartStrength = 30
                intStartDex = 20
                intStartVit = 25
                intStartEnergy = 10
                intStartLife = 55
                intStartStamina = 92
                intStartMana = 10
                decPerLevelLife = 2
                decPerLevelStam = 1
                decPerLevelMana = 1
                decPerVitLife = 4
                decPerVitStam = 1
                decPerEnergyMana = 1
            Case "Druid"
                intStartStrength = 15
                intStartDex = 20
                intStartVit = 25
                intStartEnergy = 20
                intStartLife = 55
                intStartStamina = 84
                intStartMana = 20
                decPerLevelLife = 1.5
                decPerLevelStam = 1
                decPerLevelMana = 2
                decPerVitLife = 2
                decPerVitStam = 1
                decPerEnergyMana = 2
            Case "Necromancer"
                intStartStrength = 15
                intStartDex = 25
                intStartVit = 15
                intStartEnergy = 25
                intStartLife = 45
                intStartStamina = 79
                intStartMana = 25
                decPerLevelLife = 1.5
                decPerLevelStam = 1
                decPerLevelMana = 2
                decPerVitLife = 2
                decPerVitStam = 1
                decPerEnergyMana = 2
            Case "Paladin"
                intStartStrength = 25
                intStartDex = 20
                intStartVit = 25
                intStartEnergy = 15
                intStartLife = 55
                intStartStamina = 89
                intStartMana = 15
                decPerLevelLife = 2
                decPerLevelStam = 1
                decPerLevelMana = 1.5
                decPerVitLife = 3
                decPerVitStam = 1
                decPerEnergyMana = 1.5
            Case "Sorceress"
                intStartStrength = 10
                intStartDex = 25
                intStartVit = 10
                intStartEnergy = 35
                intStartLife = 40
                intStartStamina = 74
                intStartMana = 35
                decPerLevelLife = 1
                decPerLevelStam = 1
                decPerLevelMana = 2
                decPerVitLife = 2
                decPerVitStam = 1
                decPerEnergyMana = 2
            Case Else
                intStartStrength = 0
                intStartDex = 0
                intStartVit = 0
                intStartEnergy = 0
                intStartLife = 0
                intStartStamina = 0
                intStartMana = 0
                decPerLevelLife = 0
                decPerLevelStam = 0
                decPerLevelMana = 0
                decPerVitLife = 0
                decPerVitStam = 0
                decPerEnergyMana = 0
        End Select

        'Show attributes to user
        Refresh_Data()
    End Sub

    Private Sub Level_Select_Event(sender As Object, e As EventArgs) Handles cmbLevel.SelectedIndexChanged
        'Set character level multiplier
        intCharLevel = cmbLevel.SelectedIndex

        'Show attributes to user
        Refresh_Data()
    End Sub

    Private Sub Button_Click_Events(sender As Object, e As EventArgs) Handles btnDexMinus1.Click, btnDexMinus5.Click, btnDexPlus1.Click,
        btnDexPlus5.Click, btnEnergyMinus1.Click, btnEnergyMinus5.Click, btnEnergyPlus1.Click, btnEnergyPlus5.Click, btnStrMinus1.Click,
        btnStrMinus5.Click, btnStrPlus1.Click, btnStrPlus5.Click, btnVitMinus1.Click, btnVitMinus5.Click, btnVitPlus1.Click,
        btnVitPlus5.Click

        Select Case sender.tag
            Case "str+1"
                intAddedStrength += 1
            Case "str-1"
                intAddedStrength -= 1
            Case "str+5"
                intAddedStrength += 5
            Case "str-5"
                intAddedStrength -= 5
            Case "dex+1"
                intAddedDex += 1
            Case "dex-1"
                intAddedDex -= 1
            Case "dex+5"
                intAddedDex += 5
            Case "dex-5"
                intAddedDex -= 5
            Case "vit+1"
                intAddedVit += 1
            Case "vit-1"
                intAddedVit -= 1
            Case "vit+5"
                intAddedVit += 5
            Case "vit-5"
                intAddedVit -= 5
            Case "en+1"
                intAddedEnergy += 1
            Case "en-1"
                intAddedEnergy -= 1
            Case "en+5"
                intAddedEnergy += 5
            Case "en-5"
                intAddedEnergy -= 5
        End Select

        Refresh_Data()
    End Sub

    Private Sub Plus_Buttons_Enable_Event(sender As Object, e As EventArgs) Handles lblAvailableStat.TextChanged
        'enable plus buttons only if available stats > 0
        If intAvailableStat > 0 AndAlso intAvailableStat < 5 Then
            btnDexPlus1.Enabled = True
            btnEnergyPlus1.Enabled = True
            btnStrPlus1.Enabled = True
            btnVitPlus1.Enabled = True
            btnDexPlus5.Enabled = False
            btnEnergyPlus5.Enabled = False
            btnStrPlus5.Enabled = False
            btnVitPlus5.Enabled = False
        ElseIf intAvailableStat > 4 Then
            btnDexPlus1.Enabled = True
            btnEnergyPlus1.Enabled = True
            btnStrPlus1.Enabled = True
            btnVitPlus1.Enabled = True
            btnDexPlus5.Enabled = True
            btnEnergyPlus5.Enabled = True
            btnStrPlus5.Enabled = True
            btnVitPlus5.Enabled = True
        Else
            btnDexPlus1.Enabled = False
            btnEnergyPlus1.Enabled = False
            btnStrPlus1.Enabled = False
            btnVitPlus1.Enabled = False
            btnDexPlus5.Enabled = False
            btnEnergyPlus5.Enabled = False
            btnStrPlus5.Enabled = False
            btnVitPlus5.Enabled = False
        End If
    End Sub

    Private Sub Minus_Buttons_Enable_Events(sender As Object, e As EventArgs) Handles lblAddDex.TextChanged, lblAddEnergy.TextChanged,
        lblAddStrength.TextChanged, lblAddVit.TextChanged

        'Enable minus buttons if added points > 0
        If intAddedDex > 0 AndAlso intAddedDex < 5 Then
            btnDexMinus1.Enabled = True
            btnDexMinus5.Enabled = False
        ElseIf intAddedDex > 4 Then
            btnDexMinus1.Enabled = True
            btnDexMinus5.Enabled = True
        Else
            btnDexMinus1.Enabled = False
            btnDexMinus5.Enabled = False
        End If

        If intAddedEnergy > 0 AndAlso intAddedEnergy < 5 Then
            btnEnergyMinus1.Enabled = True
            btnEnergyMinus5.Enabled = False
        ElseIf intAddedEnergy > 4 Then
            btnEnergyMinus1.Enabled = True
            btnEnergyMinus5.Enabled = True
        Else
            btnEnergyMinus1.Enabled = False
            btnEnergyMinus5.Enabled = False
        End If

        If intAddedStrength > 0 AndAlso intAddedStrength < 5 Then
            btnStrMinus1.Enabled = True
            btnStrMinus5.Enabled = False
        ElseIf intAddedStrength > 4 Then
            btnStrMinus1.Enabled = True
            btnStrMinus5.Enabled = True
        Else
            btnStrMinus1.Enabled = False
            btnStrMinus5.Enabled = False
        End If

        If intAddedVit > 0 AndAlso intAddedVit < 5 Then
            btnVitMinus1.Enabled = True
            btnVitMinus5.Enabled = False
        ElseIf intAddedVit > 4 Then
            btnVitMinus1.Enabled = True
            btnVitMinus5.Enabled = True
        Else
            btnVitMinus1.Enabled = False
            btnVitMinus5.Enabled = False
        End If
    End Sub

    Private Sub Skill_Plus_One_Events(sender As Object, e As EventArgs) Handles chkDenNorm.CheckedChanged, chkDenNight.CheckedChanged,
        chkDenHell.CheckedChanged, chkRadNorm.CheckedChanged, chkRadNight.CheckedChanged, chkRadHell.CheckedChanged

        If sender.Checked = True Then
            intAddedSkill += 1
        Else
            intAddedSkill -= 1
        End If

        Refresh_Data()
    End Sub

    Private Sub Stat_Plus_Five_Events(sender As Object, e As EventArgs) Handles chkTomeNorm.CheckedChanged, chkTomeNight.CheckedChanged,
        chkTomeHell.CheckedChanged

        If sender.Checked = True Then
            intAddedStat += 5
        Else
            intAddedStat -= 5
        End If

        Refresh_Data()
    End Sub

    Private Sub Life_Plus_Twenty_Events(sender As Object, e As EventArgs) Handles chkBirdNorm.CheckedChanged, chkBirdNight.CheckedChanged,
        chkBirdHell.CheckedChanged

        If sender.Checked = True Then
            intAddedLife += 20
        Else
            intAddedLife -= 20
        End If

        Refresh_Data()
    End Sub

    Private Sub Skill_Plus_Two_Events(sender As Object, e As EventArgs) Handles chkIzzyNorm.CheckedChanged, chkIzzyNight.CheckedChanged,
        chkIzzyHell.CheckedChanged

        If sender.Checked = True Then
            intAddedSkill += 2
        Else
            intAddedSkill -= 2
        End If

        Refresh_Data()
    End Sub
End Class
