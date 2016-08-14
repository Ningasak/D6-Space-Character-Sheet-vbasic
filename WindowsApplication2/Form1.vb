'18 dice or 54 pips Attributes
'Max 5D
'7 dice or 21 pips
'Max 3D to any skill
'greatest skill is therefore 8D


Public Class Form1
    Dim agilitySwap As New AgilitySkillPage

    Private Sub DomainUpDown1_SelectedItemChanged(sender As Object, e As EventArgs) Handles DomainUpDown1.SelectedItemChanged
        If (DomainUpDown1.Text = "Pick a race") Then
            Label1.Text = "None Selected"
        Else
            Label1.Text = DomainUpDown1.SelectedItem.ToString()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If (CheckBox1.CheckState = 1) Then
            DomainUpDown2.Show()
            Button7.Show()
            MetaphysicsToolStripMenuItem.Visible = True
        Else
            DomainUpDown2.Hide()
            Button7.Hide()
            MetaphysicsToolStripMenuItem.Visible = False
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        agilitySwap.Show()
    End Sub

    Private Sub AgilityToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AgilityToolStripMenuItem.Click
        agilitySwap.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Sheet.IncreaseStat(0)
        Label2.Text = Pip2String(Sheet.GetStat(0))
        Label9.Text = Sheet.RemainingPips
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Sheet.DecreaseStat(0)
        Label2.Text = Pip2String(Sheet.GetStat(0))
        Label9.Text = Sheet.RemainingPips
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles STRUPButton.Click
        Sheet.IncreaseStat(1)
        Label3.Text = Pip2String(Sheet.GetStat(1))
        Label9.Text = Sheet.RemainingPips
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Sheet.DecreaseStat(1)
        Label3.Text = Pip2String(Sheet.GetStat(1))
        Label9.Text = Sheet.RemainingPips
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Sheet.IncreaseStat(2)
        Label4.Text = Pip2String(Sheet.GetStat(2))
        Label9.Text = Sheet.RemainingPips
    End Sub

    Private Sub Button3_Click_2(sender As Object, e As EventArgs) Handles Button3.Click
        Sheet.DecreaseStat(2)
        Label4.Text = Pip2String(Sheet.GetStat(2))
        Label9.Text = Sheet.RemainingPips
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Sheet.IncreaseStat(3)
        Label5.Text = Pip2String(Sheet.GetStat(3))
        Label9.Text = Sheet.RemainingPips
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Sheet.DecreaseStat(3)
        Label5.Text = Pip2String(Sheet.GetStat(3))
        Label9.Text = Sheet.RemainingPips
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Sheet.IncreaseStat(4)
        Label6.Text = Pip2String(Sheet.GetStat(4))
        Label9.Text = Sheet.RemainingPips
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Sheet.DecreaseStat(4)
        Label6.Text = Pip2String(Sheet.GetStat(4))
        Label9.Text = Sheet.RemainingPips
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Sheet.IncreaseStat(5)
        Label7.Text = Pip2String(Sheet.GetStat(5))
        Label9.Text = Sheet.RemainingPips
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Sheet.DecreaseStat(5)
        Label7.Text = Pip2String(Sheet.GetStat(5))
        Label9.Text = Sheet.RemainingPips
    End Sub
End Class

