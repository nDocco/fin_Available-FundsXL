Attribute VB_Name = "Diff_Module"
Public Function MonthDiff(FromDate As Date, ToDate As Date, MonthDay As Integer) As Double
Dim Diff As Integer, DayFrom As Integer, DayTo As Integer, FromMD As Integer, ToMD As Integer, mend As Integer

    DayFrom = day(FromDate)
    'Find month day applicable to From Date
    mend = day(DateSerial(Year(FromDate), Month(FromDate) + 1, 0))
    If mend < MonthDay Then
        FromMD = mend
    Else
        FromMD = MonthDay
    End If

    DayTo = day(ToDate)
    'Find month day applicable to To Date
    mend = day(DateSerial(Year(ToDate), Month(ToDate) + 1, 0))
    If mend < MonthDay Then
        ToMD = mend
    Else
        ToMD = MonthDay
    End If

    Diff = DateDiff("m", FromDate, ToDate)
    If (DayFrom < FromMD) And (DayTo >= ToMD) Then
        Diff = Diff + 1
    ElseIf (DayFrom >= FromMD) And (DayTo < ToMD) Then
        Diff = Diff - 1
    End If
    MonthDiff = Diff
End Function
Public Function BiDiff(FromDate As Date, ToDate As Date, WkDay As Integer) As Double

    BiDiff = Int(DateDiff("ww", FromDate, ToDate, WkDay) / 2)
End Function
Public Function WeekDiff(FromDate As Date, ToDate As Date, WkDay As Integer) As Double

    WeekDiff = DateDiff("ww", FromDate, ToDate, WkDay)
End Function
Public Function RegDiff(FromDate As Date, ToDate As Date, Period As Integer) As Double
Dim Diff As Integer

    RegDiff = Int((ToDate - FromDate) / Period)
End Function
