Public Module ModCharacterSheet
    Public Sheet As New CharacterSheet

    Function Pip2String(ByVal pips As Integer)
        Dim strReturnString As String
        strReturnString = Math.Floor(pips / 3) & "D6 +" & (pips Mod 3)
        Return strReturnString
    End Function

    Public Class CharacterSheet
        Dim CharacterInformation(2) As String '0 name 1 race
        Dim Stats(7) As Integer '0 Agility 1 Strength 2 Mechanical 3 Knowledge 4 Perception 5 Technical 6 MetaPhysics
        Dim StatsBase(7) As Integer '0 Agility 1 Strength 2 Mechanical 3 Knowledge 4 Perception 5 Technical 6 MetaPhysics
        Dim Pips(3) As Integer '0 Starting Pips 1 Spent Pips 2 XP Pips
        Dim SkillMax = New Integer() {11, 4, 13, 8, 12, 12, 3} 'Number of Skills per category
        Dim ForceSensitivity As Boolean
        Dim PipsToSpendOnStats As Integer
        Dim PipsToSpendOnSkills As Integer
        Dim ASkills As New AgilitySkillPage
        Dim SSkills As New StrengthSkillPage
        Dim MSkills As New MechanicalSkillPage
        'Constructors, 
        'Default Basic Pips
        Sub New()
            PipsToSpendOnStats = 54
            PipsToSpendOnSkills = 21
            For Index As Integer = 0 To 6
                Stats(Index) = 3
                StatsBase(Index) = 3
            Next
        End Sub
        'Constructor with modifiable pips
        Sub New(ByRef StatsPips As Integer, ByRef SkillPips As Integer)
            PipsToSpendOnStats = StatsPips
            PipsToSpendOnSkills = SkillPips
            For Index As Integer = 0 To 6
                Stats(Index) = 3
                StatsBase(Index) = 3
            Next
        End Sub
        'Increases a stat by 1
        Function IncreaseStat(ByVal StatIndex)
            If (StatIndex < 0 Or StatIndex > 7) Then
                MsgBox("IncreaseStat out of range")
                Return -1 'error state
            End If
            If (Stats(StatIndex) > 20) Then
                MsgBox("Max Stat")
                Return Stats(StatIndex)
            End If
            Stats(StatIndex) += 1
            PipsToSpendOnStats -= 1
            Return Stats(StatIndex)
        End Function
        'Decreases a stat by 1 returns the value
        Function DecreaseStat(ByVal StatIndex)
            If (StatIndex < 0 Or StatIndex > 7) Then
                MsgBox("DecreaseStat out of range")
                Return -1 'error state
            End If
            If (Stats(StatIndex) < 4) Then
                MsgBox("Minimum Stat")
                Return Stats(StatIndex)
            End If
            Stats(StatIndex) -= 1
            PipsToSpendOnStats += 1
            Return Stats(StatIndex)
        End Function
        Function GetStat(ByVal StatIndex)
            If (StatIndex < 0 Or StatIndex > 7) Then
                MsgBox("GetStat out of range")
                Return -1 'error state
            End If
            Return Stats(StatIndex)
        End Function
        'Increases An Agility Skill by 1. The Skill Index 0 is Acrobatics
        Function IncreaseAgilitySkill(ByVal SkillIndex As Integer)
            Dim AdjustedValue As Integer
            If (Pips(1) > PipsToSpendOnStats) Then
                MsgBox("No more Pips Remain")
                Return ASkills.SkillValue(SkillIndex)
            End If
            If (SkillIndex < 0 Or SkillIndex > SkillMax(0)) Then
                MsgBox("Out of range index request IncreaseAgilitySkill")
                Return -1 'error
            End If
            AdjustedValue = ASkills.IncreaseSkill(SkillIndex)
            PipsToSpendOnStats -= 1
            Return AdjustedValue
        End Function
        Function RemainingPips()
            Return PipsToSpendOnStats
        End Function
        Function GetSkill(ByVal StatIndex As Integer, ByVal SkillIndex As Integer)
            If (StatIndex < 0 Or StatIndex > 7) Then
                MsgBox("Out of Range error GetSkill")
                Return -1 'error
            End If
            Select Case StatIndex
                Case 0
                    Return ASkills.SkillValue(SkillIndex)
            End Select
            Return -1 'error
        End Function
    End Class

    Public Class AgilitySkillPage
        Dim BasePips As Integer = 3 'Defaults to 1D
        Dim TotalSkills As Integer = 11
        Dim BaseIndex(TotalSkills) As Integer
        Dim AdjustedIndex(TotalSkills) As Integer

        'count for skill useage
        'Dim intAgilityIndex As Integer
        'Dim intAgilityAdjustedIndex As Integer
        'Dim intAcrobaticsIndex As Integer '0
        'Dim intAcrobaticsAdjustedIndex As Integer
        'Dim intBrawlingIndex As Integer '1
        'Dim intBrawlingAdjustedIndex As Integer
        'Dim intDodgeIndex As Integer '2
        'Dim intDodgeAdjustedIndex As Integer
        'Dim FirearmsIndex As Integer '3
        'Dim FirearmsAdjustedIndex As Integer
        'Dim Flying0GIndex As Integer '4
        'Dim Flying0GAdjustedIndex As Integer
        'Dim MeleeCombatIndex As Integer '5
        'Dim MeleeCombatAdjustedIndex As Integer
        'Dim MissileWeaponsIndex As Integer '6
        'Dim MissileWeaponsAdjustedIndex As Integer
        'Dim ThrowingIndex As Integer '7
        'Dim ThrowingAdjustedIndex As Integer
        'Dim RidingIndex As Integer '8
        'Dim RidingAdjustedIndex As Integer
        'Dim RunningIndex As Integer '9
        'Dim RunningAdjustedIndex As Integer
        'Dim SleightOfHandIndex As Integer '10
        'Dim SleightOfHandAdjustedIndex As Integer
        'Dim SpentPipsBase As Integer
        'Dim SpentPipsAdjusted As Integer


        'Dim SkillListArray(11) As String

        'Dim strAcrobatics As String '0
        'Dim strBrawling As String
        'Dim strDodge As String
        'Dim strFirearms As String
        'Dim strFlying0G As String
        'Dim strMeleeCombat As String
        'Dim strMissileWeapons As String
        'Dim strRiding As String
        'Dim strRunning As String
        'Dim strSleightOfHand As String
        'Dim strThrowing As String '10
        Sub New()
            BasePips = 3
            SetBaseSkills()
        End Sub
        Sub New(ByVal pips As Integer)
            BasePips = pips
            SetBaseSkills()
        End Sub
        Sub SetBaseSkills()
            For count = 0 To 10
                BaseIndex(count) = BasePips
                AdjustedIndex(count) = BasePips
            Next
        End Sub
        Function IncreaseSkill(ByVal ReferenceIndex As Integer)
            If (ReferenceIndex < 0 Or ReferenceIndex > TotalSkills) Then
                MsgBox("Out of range index request IncreaseAgilitySkill")
                Return -1 'error
            End If
            AdjustedIndex(ReferenceIndex) += 1
            Return AdjustedIndex(ReferenceIndex)
        End Function
        Function SkillValue(ByVal ReferenceIndex)
            Return AdjustedIndex(ReferenceIndex)
        End Function


        Function ModifyStatByBasePips(ByRef ArrayToChange As Integer, ByRef TrueIsPositive As Boolean)
            If (ArrayToChange > TotalSkills) Then
                MsgBox("Out of Range Exception Agility Skills list ModifyStatsByBasePips")
                Return 0
            End If
            If (TrueIsPositive) Then
                AdjustedIndex(ArrayToChange) += 1
                Return 1
            Else
                AdjustedIndex(ArrayToChange) -= 1
                Return -1
            End If
        End Function

        'Public Sub OR when it launches it should do this...initializeList() sets the index to the base pips
        'have each adjustment adjust the total spentPips
    End Class
    Public Class StrengthSkillPage

    End Class
    Public Class MechanicalSkillPage

    End Class
    Public Class KnowledgeSkillPage

    End Class
    Public Class PerceptionSkillPage

    End Class
    Public Class TechnicalSkillPage

    End Class
    Public Class MetaphysicsSkillPage

    End Class
End Module


'Select Case ArrayToChange
'Case 0
'strAcrobatics = PipToString(AdjustedIndex(0))
'Case 1
'strBrawling = AdjustedIndex(1)
'Case 2
'strDodge = AdjustedIndex(2)
'Case 3
'strFirearms = AdjustedIndex(3)
'Case 4
'strFlying0G = AdjustedIndex(4)
'Case 5
'strMeleeCombat = AdjustedIndex(5)
'Case 6
'strMissileWeapons = AdjustedIndex(6)
'Case 7
'strRiding = AdjustedIndex(7)
'Case 8
'strRunning = AdjustedIndex(8)
'Case 9
'strSleightOfHand = AdjustedIndex(9)
'Case 10
'strThrowing = AdjustedIndex(10)
'End Select

'Sub ArrayToIndividual()
'    strAcrobatics = SkillListArray(0)
'    strBrawling = SkillListArray(1)
'    strDodge = SkillListArray(2)
'    strFirearms = SkillListArray(3)
'    strFlying0G = SkillListArray(4)
'    strMeleeCombat = SkillListArray(5)
'    strMissileWeapons = SkillListArray(6)
'    strRiding = SkillListArray(7)
'    strRunning = SkillListArray(8)
'    strSleightOfHand = SkillListArray(9)
'    strThrowing = SkillListArray(10)
'End Sub

'Sub setBaseSkills()
'    For count = 1 To 11
'        BaseIndex(count) = SpentPipsBase
'        AdjustedIndex(count) = SpentPipsBase
'        SkillListArray(count - 1) = Pip2String(SpentPipsBase)
'    Next
'    ArrayToIndividual()
'End Sub
'Sub ModifyStat(ByRef ArrayToChange As Integer, ByRef AmountToChange As Integer)
'    AdjustedIndex(ArrayToChange) += AmountToChange
'    Select Case ArrayToChange
'        Case 0
'            strAcrobatics = AdjustedIndex(0)
'        Case 1
'            strBrawling = AdjustedIndex(1)
'        Case 2
'            strDodge = AdjustedIndex(2)
'        Case 3
'            strFirearms = AdjustedIndex(3)
'        Case 4
'            strFlying0G = AdjustedIndex(4)
'        Case 5
'            strMeleeCombat = AdjustedIndex(5)
'        Case 6
'            strMissileWeapons = AdjustedIndex(6)
'        Case 7
'            strRiding = AdjustedIndex(7)
'        Case 8
'            strRunning = AdjustedIndex(8)
'        Case 9
'            strSleightOfHand = AdjustedIndex(9)
'        Case 10
'            strThrowing = AdjustedIndex(10)
'    End Select
'End Sub

'Public Class CharacterSheet
'    Dim CharacterInformation(2) As String '0 name 1 race
'    Dim Stats(7) As Integer '0 Agility 1 Strength 2 Mechanical 3 Knowledge 4 Perception 5 Technical 6 MetaPhysics
'    Dim Pips(3) As Integer '0 Starting Pips 1 Spent Pips 2 XP Pips
'    Dim ForceSensitivity As Boolean

'    Dim PipsToSpendOnStats As Integer = 54
'    Dim ASkills As New AgilitySkillPage
'    Dim SSkills As New StrengthSkillPage
'    Dim MSkills As New MechanicalSkillPage

'    Dim CharacterName As String
'    Dim CharacterRace As String
'    Dim Agility As String
'    Dim Strength As String
'    Dim Mechanical As String
'    Dim Knowledge As String
'    Dim Perception As String
'    Dim Technical As String
'    Dim Metaphysics As String

'End Class