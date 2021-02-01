Attribute VB_Name = "Proj_Module"
Public Function ProjAnn(Amount As Double, First As Date, Final As Date, Current As Date, Projected As Date) As Double
Dim Previous As Double
Dim SaveDate As Date, Due As Date, DueSave As Date

    'Find Date for projected cycle / final payment
    Due = NextDate(Projected, First, Final, "A")
    DueSave = DateAdd("yyyy", -1, Due)

    'Have payments completed?
    If (Final > 0) Then
        If Current >= Due Then
            ProjAnn = 0
            Exit Function
        ElseIf Projected > Final Then
            DueSave = Due
        End If
    End If

    'Calculate save date for first payment
    SaveDate = DateAdd("yyyy", -1, First)

    'Have payments started?
    If Projected < SaveDate Then
        ProjAnn = 0
        Exit Function
    ElseIf Current > First Then
        'Update save date to current cycle
        SaveDate = NextDate(Current, First, Final, "A")
        SaveDate = DateAdd("yyyy", -1, SaveDate)
    End If

    'Cycles Completed
    Previous = DateDiff("yyyy", SaveDate, DueSave)

    'Calculate % of projected month completed
    ProjAnn = CalcSave(Amount, Projected, DueSave, Due, Previous)
End Function
Public Function ProjMon(Amount As Double, MonthDay As Integer, First As Date, Final As Date, Current As Date, Projected As Date) As Double
Dim Previous As Double
Dim SaveDate As Date, DueSave As Date, Due As Date

    'Find Date for projected cycle / final payment
    Due = NextDate(Projected, First, Final, "M", MonthDay)
    DueSave = MonthAdd(Due, -1, MonthDay)

    'Have payments completed?
    If (Final > 0) Then
        If Current >= Due Then
            ProjMon = 0
            Exit Function
        ElseIf Projected > Final Then
            DueSave = Due
        End If
    End If

    'Calculate save date for first payment
    SaveDate = MonthAdd(First, -1, MonthDay)

    'Have payments started?
    If Projected < SaveDate Then
        ProjMon = 0
        Exit Function
    ElseIf Current > First Then
        'Update save date to current cycle
        SaveDate = NextDate(Current, First, Final, "M", MonthDay)
        SaveDate = MonthAdd(SaveDate, -1, MonthDay)
    End If

    'Cycles Completed
    Previous = MonthDiff(SaveDate, DueSave, MonthDay)

    'Calculate % of projected month completed
    ProjMon = CalcSave(Amount, Projected, DueSave, Due, Previous)
End Function
Public Function ProjBi(Amount As Double, WkDay As Integer, First As Date, Final As Date, Current As Date, Projected As Date) As Double
Dim Previous As Double
Dim SaveDate As Date, Due As Date, DueSave As Date

    'Find Date for projected cycle / final payment
    Due = NextDate(Projected, First, Final, "B", WkDay)
    DueSave = DateAdd("ww", -2, Due)

    'Have payments completed?
    If (Final > 0) Then
        If Current >= Due Then
            ProjBi = 0
            Exit Function
        ElseIf Projected > Final Then
            DueSave = Due
        End If
    End If

    'Calculate save date for first payment
    SaveDate = NextDOW(First, WkDay)
    SaveDate = DateAdd("ww", -3, First)

    'Have payments started?
    If Projected < SaveDate Then
        ProjBi = 0
        Exit Function
    ElseIf Current > First Then
        'Update save date to current cycle
        SaveDate = NextDate(Current, First, Final, "B", WkDay)
        SaveDate = DateAdd("ww", -2, SaveDate)
    End If

    'Cycles Completed
    Previous = BiDiff(SaveDate, DueSave, WkDay)
'    MsgBox "Previous payments: " & Previous
    
    'Calculate % of projected month completed
    ProjBi = CalcSave(Amount, Projected, DueSave, Due, Previous)
End Function
Public Function ProjWk(Amount As Double, WkDay As Integer, First As Date, Final As Date, Current As Date, Projected As Date) As Double
Dim Previous As Double
Dim SaveDate As Date, Due As Date, DueSave As Date

    'Find Date for projected cycle / final payment
    Due = NextDate(Projected, First, Final, "W", WkDay)
    DueSave = DateAdd("ww", -1, Due)

    'Have payments completed?
    If (Final > 0) Then
        If Current >= Due Then
            ProjWk = 0
            Exit Function
        ElseIf Projected > Final Then
            DueSave = Due
        End If
    End If

    'Calculate save date for first payment
    SaveDate = NextDOW(First, WkDay)
    SaveDate = DateAdd("ww", -1, First)

    'Have payments started?
    If Projected < SaveDate Then
        ProjWk = 0
        Exit Function
    ElseIf Current > First Then
        'Update save date to current cycle
        SaveDate = NextDate(Current, First, Final, "W", WkDay)
        SaveDate = DateAdd("ww", -1, SaveDate)
    End If

    'Cycles Completed
    Previous = WeekDiff(SaveDate, DueSave, WkDay)
    
    'Calculate % of projected month completed
    ProjWk = CalcSave(Amount, Projected, DueSave, Due, Previous)
End Function
Public Function ProjReg(Amount As Double, Period As Integer, First As Date, Final As Date, Current As Date, Projected As Date) As Double
Dim Previous As Double
Dim RegPer As String
Dim SaveDate As Date, Due As Date, DueSave As Date
    
    RegPer = CStr(Period)

    'Find Date for projected cycle / final payment
    Due = NextDate(Projected, First, Final, RegPer)
    DueSave = DateAdd("d", -Period, Due)

    'Have payments completed?
    If (Final > 0) Then
        If Current >= Due Then
            ProjReg = 0
            Exit Function
        ElseIf Projected > Final Then
            DueSave = Due
        End If
    End If

    'Calculate save date for first payment
    SaveDate = DateAdd("d", -Period, First)

    'Have payments started?
    If Projected < SaveDate Then
        ProjReg = 0
        Exit Function
    ElseIf Current > First Then
        'Update save date to current cycle
        SaveDate = NextDate(Current, First, Final, RegPer)
        SaveDate = DateAdd("d", -Period, SaveDate)
    End If

    'Cycles Completed
    Previous = RegDiff(SaveDate, DueSave, Period)
    
    'Calculate % of projected month completed
    ProjReg = CalcSave(Amount, Projected, DueSave, Due, Previous)
End Function

