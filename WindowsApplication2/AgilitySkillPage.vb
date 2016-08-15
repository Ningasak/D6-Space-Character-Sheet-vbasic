Public Class AgilitySkillPage

    Private Sub DomainUpDown1_SelectedItemChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ButtonAcrobaticPlus_Click(sender As Object, e As EventArgs) Handles ButtonAcrobaticPlus.Click
        Sheet.IncreaseAgilitySkill(0)
        UpdatePipsCounters()
        AcrobaticsRollLabel.Text = Pip2String(Sheet.GetSkill(0, 0))
    End Sub
    Private Sub UpdatePipsCounters()
        Dim Pip As Integer
        Pip = Sheet.RemainingPips
        Label2.Text = (54 - Pip)
        Label4.Text = Pip
    End Sub

    Private Sub AgilitySkillPage_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        AgilityTotalDice.Text = Pip2String(Sheet.GetStat(0))
        'Sub that takes and compares the old value to the new value
        'if there is a change -/+ then reset the base and add the difference to skill total
        AcrobaticsRollLabel.Text = Pip2String(Sheet.GetSkill(0, 0))
        BrawlingRollLabel.Text = Pip2String(Sheet.GetSkill(0, 1))
        DodgeRollLabel.Text = Pip2String(Sheet.GetSkill(0, 2))
        FirearmsRollLabel.Text = Pip2String(Sheet.GetSkill(0, 3))
        FlyingRollLabel.Text = Pip2String(Sheet.GetSkill(0, 4))
        MeleeCombatRollLabel.Text = Pip2String(Sheet.GetSkill(0, 5))
        MissleWeaponRollLabel.Text = Pip2String(Sheet.GetSkill(0, 6))
        RidingRollLabel.Text = Pip2String(Sheet.GetSkill(0, 7))
        RunningRollLabel.Text = Pip2String(Sheet.GetSkill(0, 8))
        SleightOfHandRollLabel.Text = Pip2String(Sheet.GetSkill(0, 9))
        ThrowingRollLabel.Text = Pip2String(Sheet.GetSkill(0, 10))
    End Sub

    Private Sub AgilitySkillPage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim PresetSkills As New Collection 'this whole thing should be a each next loop that loads from a file. Remake it dummy.
        PresetSkills.Add(AcrobaticsRollLabel.Text)
        PresetSkills.Add(BrawlingRollLabel.Text)
        PresetSkills.Add(DodgeRollLabel.Text)
        PresetSkills.Add(FirearmsRollLabel.Text)
        PresetSkills.Add(FlyingRollLabel.Text)
        PresetSkills.Add(MeleeCombatRollLabel.Text)
        PresetSkills.Add(MissleWeaponRollLabel)
        PresetSkills.Add(RidingRollLabel.Text)
        PresetSkills.Add(RunningRollLabel)
        PresetSkills.Add(SleightOfHandRollLabel)
        PresetSkills.Add(ThrowingRollLabel)
        PresetSkills = Pip2String(Sheet.GetStat(0))
    End Sub
End Class