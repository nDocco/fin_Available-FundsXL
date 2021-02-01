Attribute VB_Name = "Global_Module"
Public Function NextDate(Focus As Date, First As Date, Final As Date, Period As String, Optional MonthDay As Integer) As Date
Dim ATest As String, MTest As String, BTest As String, WTest As String, RTest As String
Dim adjustment As Double, interval As Double
Dim WkDay As Integer, RegPer As Integer
Dim Template As Date, Start As Date

Start = First

ATest = "A"
MTest = "M"
BTest = "B"
WTest = "W"
RTest = "R"

'Have Payments started?
If Focus < Start Then
    NextDate = Start
'     MsgBox "Payments have not started"
    Exit Function
End If

'Test for Annual
If LCase(Period) = LCase(ATest) Then
' MsgBox "Annual Payments"

    'No of years from Start
    interval = Year(Focus) - Year(Start)
    NextDate = DateAdd("yyyy", interval, Start)
'     MsgBox "This years payment is " & NextDate

    'If pay date has passed next date
    If NextDate <= Focus Then
        NextDate = DateAdd("yyyy", 1, NextDate)
'         MsgBox "Date Passed. Next Payment is " & NextDate
    End If
    
    'Has final date been passed?
    If (Final > 0) And (NextDate > Final) Then
'         MsgBox "Final payment passed"
        interval = DateDiff("yyyy", NextDate, Final)
'         MsgBox "Final payment " & interval & " years ago"
        NextDate = DateAdd("yyyy", interval, NextDate)
'         MsgBox "Final Payment now " & NextDate
    End If

'Test for Monthly
ElseIf LCase(Period) = LCase(MTest) Then
' MsgBox "Monthly Payments"

    'Was start parameter passed?
    If Start = 0 Then
        If IsMissing(MonthDay) Then
            MsgBox "Error.  Must provide start date or day of month for payments."
            Exit Function
        End If
        Start = DateSerial(Year(Focus), 1, MonthDay)
'         MsgBox "Missing start date is now: " & Start
    End If
    
    'Provide default for MonthDay if missing
    If IsMissing(MonthDay) = True Then MonthDay = day(Start)
    
    'No of months from Start
    interval = DateDiff("m", Start, Focus)
    NextDate = MonthAdd(Start, interval, MonthDay)
'     MsgBox "This months payment is on " & NextDate

    'Check if pay date is after this month's payment
    If NextDate <= Focus Then
        NextDate = MonthAdd(NextDate, 1, MonthDay)
'         MsgBox "Date Passed. Next Payment is " & NextDate
    End If
    
    'Has final date been passed?
    If (Final > 0) And (NextDate > Final) Then
'         MsgBox "Final payment passed"
        interval = DateDiff("m", NextDate, Final)
'         MsgBox "Final interval is " & interval & " months ago"
        NextDate = MonthAdd(NextDate, interval, MonthDay)
'         MsgBox "Final payment now " & NextDate
    End If
    
'Test for Bi-Weekly
ElseIf LCase(Period) = LCase(BTest) Then
' MsgBox "Bi-Weekly Payments"
    'Test if Week Day parameter passed?
    If (MonthDay >= 0) And (MonthDay <= 7) Then
        WkDay = MonthDay
    Else
        WkDay = 0
    End If

    'Find next instance of week day
    NextDate = NextDOW(Focus, WkDay)
    
    'Check if odd or even number of weeks since start
    interval = WeekDiff(Start, NextDate, WkDay) - (BiDiff(Start, NextDate, WkDay) * 2)
    NextDate = DateAdd("ww", interval, NextDate)
    
    'Has final date been passed?
    If (Final > 0) And (NextDate > Final) Then
        interval = BiDiff(NextDate, Final, WkDay) * 2
        NextDate = DateAdd("ww", interval, NextDate)
    End If

'Test for Weekly
ElseIf LCase(Period) = LCase(WTest) Then
' MsgBox "Weekly Payments"
    'Test if Week Day parameter passed?
    If (MonthDay >= 0) And (MonthDay <= 7) Then
        WkDay = MonthDay
    Else
        WkDay = 0
    End If

    'Find next instance of week day
    NextDate = NextDOW(Focus, WkDay)
    
    'Has final date been passed?
    If (Final > 0) And (NextDate > Final) Then
        interval = WeekDiff(NextDate, Final, WkDay)
        NextDate = DateAdd("ww", interval, NextDate)
    End If
    
'Test for Regular
ElseIf IsNumeric(Period) Then
' MsgBox "Regular Payments"
    RegPer = Val(Period)
'     MsgBox "Period is " & RegPer
        
    'Find next instance of period
    interval = RegDiff(Start, Focus, RegPer) + 1
    adjustment = interval * RegPer
    NextDate = DateAdd("d", adjustment, Start)
    
    'Has final date been passed?
    If (Final > 0) And (NextDate > Final) Then
        interval = RegDiff(Start, Final, RegPer)
        adjustment = interval * RegPer
        NextDate = DateAdd("d", adjustment, Start)
    End If
Else
    MsgBox "Period was wrong"
End If
End Function
Public Function MonthAdd(Date1 As Date, adjustment As Double, MonthDay As Integer) As Date
Dim Template As Date, Temp As Date

    Temp = DateAdd("m", adjustment, Date1)
    
    adjustment = Month(Temp) - 1
    Template = DateSerial(Year(Temp), 1, MonthDay)
    
    MonthAdd = DateAdd("m", adjustment, Template)
End Function
Public Function NextDOW(Date1 As Date, DayWeek As Integer)
Dim adjustment As Double

    adjustment = 8 - WeekDay(Date1, DayWeek)
    NextDOW = Date1 + adjustment
End Function
Public Function CalcSave(Amount As Double, Focus As Date, SaveDate, Due As Date, Optional Previous As Double) As Double
Dim DaysPassed As Integer, TotalPeriod As Integer

If IsMissing(Previous) Then
    Previous = 0
End If

'Have savings started?
    If (Focus < SaveDate) Or (Focus > Due) Then
        DaysPassed = 0
    ElseIf Focus = Due Then
        CalcSave = 0
        Exit Function
    Else
        DaysPassed = Focus - SaveDate + 1
    End If

'Calculate % of saving days completed
    TotalPeriod = Due - SaveDate
'     MsgBox DaysPassed & " Days Passed for Period of " & TotalPeriod
    If DaysPassed < 1 Then
        CalcSave = (Amount * Previous)
    Else
        CalcSave = (DaysPassed / TotalPeriod * Amount) + (Amount * Previous)
    End If
End Function
