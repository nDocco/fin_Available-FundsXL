Attribute VB_Name = "Term_Module"
Public Function Annual(Amount As Double, Current As Date, Due As Date) As Double
Dim DaysPassed As Integer, TotalPeriod As Integer
Dim SaveDate As Date

    SaveDate = DateAdd("yyyy", -1, Due)
    Annual = CalcSave(Amount, Current, SaveDate, Due)
End Function
Public Function Monthly(Amount As Double, MonthDay As Integer, Current As Date, Due As Date) As Double
Dim SaveDate As Date

    SaveDate = MonthAdd(Due, -1, MonthDay)
    Monthly = CalcSave(Amount, Current, SaveDate, Due)
End Function
Public Function BiWeekly(Amount As Double, WkDay As Integer, Current As Date, Due As Date) As Double
Dim DaysPassed As Integer, TotalPeriod As Integer
Dim SaveDate As Date

    SaveDate = DateAdd("ww", -2, Due)
    BiWeekly = CalcSave(Amount, Current, SaveDate, Due)
End Function
Public Function Weekly(Amount As Double, WkDay As Integer, Current As Date, Due As Date) As Double
Dim DaysPassed As Integer, TotalPeriod As Integer
Dim SaveDate As Date

    SaveDate = DateAdd("ww", -1, Due)
    Weekly = CalcSave(Amount, Current, SaveDate, Due)
End Function
Public Function Regular(Amount As Double, Period As Integer, Current As Date, Due As Date) As Double
Dim DaysPassed As Integer, TotalPeriod As Integer
Dim SaveDate As Date

    SaveDate = DateAdd("d", -Period, Due)
    Regular = CalcSave(Amount, Current, SaveDate, Due)
End Function

