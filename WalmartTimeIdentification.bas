Attribute VB_Name = "WalmartTimeIdentification"
Option Explicit
Option Private Module

'OWNER: SICHENG (CHARLES) ZHOU
'CONTACT: CHARLES.ZHOU@DSCLOGISTICS.COM

'method to get the walmart time range
Public Sub getWalmartTimeWeek(ByVal dateInWeek As Date, ByRef firstDate As Date, ByRef lastDate As Date)
    
    If DateValue(dateInWeek) = DateValue(Common.LastDayInWeek(dateInWeek)) Then
        
        firstDate = DateValue(dateInWeek)
        lastDate = DateValue(dateInWeek + 6)
        
    Else
        
        firstDate = DateValue(Common.FirstDayInWeek(dateInWeek) - 1)
        lastDate = DateValue(Common.LastDayInWeek(dateInWeek) - 1)
    
    End If
    
End Sub


'method to get all other time range
Public Sub get_other_time_week(ByVal dateInWeek As Date, ByRef firstDate As Date, ByRef lastDate As Date)
    
    firstDate = DateValue(Common.FirstDayInWeek(dateInWeek))
    lastDate = DateValue(Common.LastDayInWeek(dateInWeek))
    
End Sub


'method to get the walmart month range
Public Sub getWalmartTimeMonth(ByVal dateInWeek As Date, ByRef firstDate As Date, ByRef lastDate As Date)
    
    Dim firstDateCurrentMonth As Date: firstDateCurrentMonth = Common.FirstDayInMonth(dateInWeek)
    Dim firstDateNextMonth As Date: firstDateNextMonth = Common.LastDayInMonth(dateInWeek) + 1
    Dim firstDateNext2Month As Date: firstDateNext2Month = Common.LastDayInMonth(firstDateNextMonth) + 1
    
    Dim walmartFirstDateInCurrentCalendarMonth As Date
    Dim walmartFirstDateInNextCalendarMonth As Date
    Dim walmartFirstDateInNext2CalendarMonth As Date
    
    Dim bufferLastDate As Date
    
    Call getWalmartTimeWeek(firstDateCurrentMonth, walmartFirstDateInCurrentCalendarMonth, bufferLastDate)
    Call getWalmartTimeWeek(firstDateNextMonth, walmartFirstDateInNextCalendarMonth, bufferLastDate)
    Call getWalmartTimeWeek(firstDateNext2Month, walmartFirstDateInNext2CalendarMonth, bufferLastDate)
    
    If dateInWeek >= walmartFirstDateInNextCalendarMonth Then
        
        firstDate = walmartFirstDateInNextCalendarMonth
        lastDate = walmartFirstDateInNext2CalendarMonth - 1
        
    Else
        
        firstDate = walmartFirstDateInCurrentCalendarMonth
        lastDate = walmartFirstDateInNextCalendarMonth - 1
        
    End If
    
    
End Sub


'method to get all other time range
Public Sub get_other_time_month(ByVal dateInWeek As Date, ByRef firstDate As Date, ByRef lastDate As Date)
    
    firstDate = Common.FirstDayInMonth(dateInWeek)
    lastDate = Common.LastDayInMonth(dateInWeek)
    
End Sub


'method to get the walmart time range
Public Sub getWalmartTimeYear(ByVal dateInWeek As Date, ByRef firstDate As Date, ByRef lastDate As Date)
    
    Dim current_year As Integer: current_year = year(dateInWeek)
    Dim saturday_day As Integer: saturday_day = 7
    
    Dim current_firstDate_q1 As Date
    Dim next_firstDate_q1 As Date
    Dim lastDate_q4 As Date
    
    'to get the first date of this year's q1
    Dim current_feb_firstDate As Date: current_feb_firstDate = DateValue("2/1/" & current_year)
    Dim next_feb_firstDate As Date: next_feb_firstDate = DateValue("2/1/" & current_year + 1)
    
    Dim current_feb_first_day As Integer: current_feb_first_day = Weekday(current_feb_firstDate)
    Dim next_feb_first_day As Integer: next_feb_first_day = Weekday(next_feb_firstDate)
    
    Dim current_feb_first_week_firstDate As Date: current_feb_first_week_firstDate = Common.FirstDayInWeek(current_feb_firstDate)
    Dim next_feb_first_week_firstDate As Date: next_feb_first_week_firstDate = Common.FirstDayInWeek(next_feb_firstDate)
    
    'get 1st date of Wal-Mart Q1 current year
    If current_feb_first_day = saturday_day Then
        current_firstDate_q1 = current_feb_firstDate
    Else
        current_firstDate_q1 = current_feb_first_week_firstDate - 1
    End If
    
    'if today's date is before the first date of Q1 in this year, then today is still in last year's Q4
    If dateInWeek < current_firstDate_q1 Then current_year = current_year - 1
    
    current_feb_firstDate = DateValue("2/1/" & current_year)
    next_feb_firstDate = DateValue("2/1/" & current_year + 1)
    
    current_feb_first_day = Weekday(current_feb_firstDate)
    next_feb_first_day = Weekday(next_feb_firstDate)
    
    current_feb_first_week_firstDate = Common.FirstDayInWeek(current_feb_firstDate)
    next_feb_first_week_firstDate = Common.FirstDayInWeek(next_feb_firstDate)
    
    
    'get 1st date of Wal-Mart Q1 current year
    If current_feb_first_day = saturday_day Then
        current_firstDate_q1 = current_feb_firstDate
    Else
        current_firstDate_q1 = current_feb_first_week_firstDate - 1
    End If
    
    'get 1st date of Wal-Mart Q1 next year
    If next_feb_first_day = saturday_day Then
        next_firstDate_q1 = next_feb_firstDate
    Else
        next_firstDate_q1 = next_feb_first_week_firstDate - 1
    End If

    
    'get the last date of a quarter
    lastDate_q4 = next_firstDate_q1 - 1
    
        
    firstDate = current_firstDate_q1
    lastDate = lastDate_q4
    
End Sub


'method to get all other time range
Public Sub getOtherTimeYear(ByVal dateInWeek As Date, ByRef firstDate As Date, ByRef lastDate As Date)
    
    Dim current_year As Integer: current_year = year(dateInWeek)
    
    firstDate = Common.FirstDayInYear(dateInWeek)
    lastDate = Common.LastDayInYear(dateInWeek)
    
End Sub



'method to get the walmart time range
Public Sub getWalmartTimeQuarter(ByVal dateInWeek As Date, ByRef firstDate As Date, ByRef lastDate As Date)
    
    Dim current_year As Integer: current_year = year(dateInWeek)
    Dim saturday_day As Integer: saturday_day = 7
    
    Dim current_firstDate_q1 As Date
    
    'to get the first date of this year's q1
    Dim current_feb_firstDate As Date: current_feb_firstDate = DateValue("2/1/" & current_year)
    Dim next_feb_firstDate As Date: next_feb_firstDate = DateValue("2/1/" & current_year + 1)
    
    Dim current_feb_first_day As Integer: current_feb_first_day = Weekday(current_feb_firstDate)
    Dim next_feb_first_day As Integer: next_feb_first_day = Weekday(next_feb_firstDate)
    
    Dim current_feb_first_week_firstDate As Date: current_feb_first_week_firstDate = Common.FirstDayInWeek(current_feb_firstDate)
    Dim next_feb_first_week_firstDate As Date: next_feb_first_week_firstDate = Common.FirstDayInWeek(next_feb_firstDate)
    
    'get 1st date of Wal-Mart Q1 current year
    If current_feb_first_day = saturday_day Then
        current_firstDate_q1 = current_feb_firstDate
    Else
        current_firstDate_q1 = current_feb_first_week_firstDate - 1
    End If
    
    'if today's date is before the first date of Q1 in this year, then today is still in last year's Q4
    If dateInWeek < current_firstDate_q1 Then current_year = current_year - 1
    
    current_feb_firstDate = DateValue("2/1/" & current_year)
    next_feb_firstDate = DateValue("2/1/" & current_year + 1)
    Dim may_firstDate As Date: may_firstDate = DateValue("5/1/" & current_year)
    Dim aug_firstDate As Date: aug_firstDate = DateValue("8/1/" & current_year)
    Dim nov_firstDate As Date: nov_firstDate = DateValue("11/1/" & current_year)
    
    current_feb_first_day = Weekday(current_feb_firstDate)
    next_feb_first_day = Weekday(next_feb_firstDate)
    Dim may_first_day As Integer: may_first_day = Weekday(may_firstDate)
    Dim aug_first_day As Integer: aug_first_day = Weekday(aug_firstDate)
    Dim nov_first_day As Integer: nov_first_day = Weekday(nov_firstDate)
    
    current_feb_first_week_firstDate = Common.FirstDayInWeek(current_feb_firstDate)
    next_feb_first_week_firstDate = Common.FirstDayInWeek(next_feb_firstDate)
    Dim may_first_week_firstDate As Date: may_first_week_firstDate = Common.FirstDayInWeek(may_firstDate)
    Dim aug_first_week_firstDate As Date: aug_first_week_firstDate = Common.FirstDayInWeek(aug_firstDate)
    Dim nov_first_week_firstDate As Date: nov_first_week_firstDate = Common.FirstDayInWeek(nov_firstDate)
    
    Dim next_firstDate_q1 As Date
    Dim lastDate_q1 As Date
    Dim firstDate_q2 As Date
    Dim lastDate_q2 As Date
    Dim firstDate_q3 As Date
    Dim lastDate_q3 As Date
    Dim firstDate_q4 As Date
    Dim lastDate_q4 As Date

    
    'get 1st date of Wal-Mart Q1 current year
    If current_feb_first_day = saturday_day Then
        current_firstDate_q1 = current_feb_firstDate
    Else
        current_firstDate_q1 = current_feb_first_week_firstDate - 1
    End If
    
    'get 1st date of Wal-Mart Q1 next year
    If next_feb_first_day = saturday_day Then
        next_firstDate_q1 = next_feb_firstDate
    Else
        next_firstDate_q1 = next_feb_first_week_firstDate - 1
    End If
    
    'get 1st date of Wal-Mart Q2
    If may_first_day = saturday_day Then
        firstDate_q2 = may_firstDate
    Else
        firstDate_q2 = may_first_week_firstDate - 1
    End If
    
    'get 1st date of Wal-Mart Q3
    If aug_first_day = saturday_day Then
        firstDate_q3 = aug_firstDate
    Else
        firstDate_q3 = aug_first_week_firstDate - 1
    End If
    
    'get 1st date of Wal-Mart Q4
    If nov_first_day = saturday_day Then
        firstDate_q4 = nov_firstDate
    Else
        firstDate_q4 = nov_first_week_firstDate - 1
    End If
    
    'get the last date of a quarter
    lastDate_q1 = firstDate_q2 - 1
    lastDate_q2 = firstDate_q3 - 1
    lastDate_q3 = firstDate_q4 - 1
    lastDate_q4 = next_firstDate_q1 - 1
    


    
    'Wal-Mart Q1
    If dateInWeek >= current_firstDate_q1 And dateInWeek <= lastDate_q1 Then
        
        firstDate = current_firstDate_q1
        lastDate = lastDate_q1
    
    'Wal-Mart Q2
    ElseIf dateInWeek >= firstDate_q2 And dateInWeek <= lastDate_q2 Then
        
        firstDate = firstDate_q2
        lastDate = lastDate_q2
        
    'Wal-Mart Q3
    ElseIf dateInWeek >= firstDate_q3 And dateInWeek <= lastDate_q3 Then
        
        firstDate = firstDate_q3
        lastDate = lastDate_q3
    
    'Wal-Mart current Q4
    ElseIf dateInWeek >= firstDate_q4 And dateInWeek <= lastDate_q4 Then
        
        firstDate = firstDate_q4
        lastDate = lastDate_q4
    
    End If
    
End Sub


'method to get all other time range
Public Sub getOtherTimeQuarter(ByVal dateInWeek As Date, ByRef firstDate As Date, ByRef lastDate As Date)
    
    Dim firstDate_q1 As Date: firstDate_q1 = DateValue("1/1/" & year(dateInWeek))
    Dim lastDate_q1 As Date: lastDate_q1 = DateValue("4/1/" & year(dateInWeek)) - 1
    Dim firstDate_q2 As Date: firstDate_q2 = DateValue("4/1/" & year(dateInWeek))
    Dim lastDate_q2 As Date: lastDate_q2 = DateValue("7/1/" & year(dateInWeek)) - 1
    Dim firstDate_q3 As Date: firstDate_q3 = DateValue("7/1/" & year(dateInWeek))
    Dim lastDate_q3 As Date: lastDate_q3 = DateValue("10/1/" & year(dateInWeek)) - 1
    Dim firstDate_q4 As Date: firstDate_q4 = DateValue("10/1/" & year(dateInWeek))
    Dim lastDate_q4 As Date: lastDate_q4 = DateValue("12/31/" & year(dateInWeek))
    
    'All other customers Q1
    If dateInWeek >= firstDate_q1 And dateInWeek <= lastDate_q1 Then
        
        firstDate = firstDate_q1
        lastDate = lastDate_q1
    
    'All other customers Q2
    ElseIf dateInWeek >= firstDate_q2 And dateInWeek <= lastDate_q2 Then
        
        firstDate = firstDate_q2
        lastDate = lastDate_q2
        
    'All other customers Q3
    ElseIf dateInWeek >= firstDate_q3 And dateInWeek <= lastDate_q3 Then
        
        firstDate = firstDate_q3
        lastDate = lastDate_q3
    
    'All other customers Q4
    ElseIf dateInWeek >= firstDate_q4 And dateInWeek <= lastDate_q4 Then
        
        firstDate = firstDate_q4
        lastDate = lastDate_q4
    
    End If
    
End Sub



'get the quarter number through a given date
Public Function getCurrentQuarterNumber(ByVal isWalmart As Boolean, ByVal dateInWeek As Date) As Integer
    
    Dim current_year As Integer: current_year = year(dateInWeek)
    
    Dim current_feb_firstDate As Date: current_feb_firstDate = DateValue("2/1/" & current_year)
    Dim next_feb_firstDate As Date: next_feb_firstDate = DateValue("2/1/" & current_year + 1)
    Dim may_firstDate As Date: may_firstDate = DateValue("5/1/" & current_year)
    Dim aug_firstDate As Date: aug_firstDate = DateValue("8/1/" & current_year)
    Dim nov_firstDate As Date: nov_firstDate = DateValue("11/1/" & current_year)
    
    Dim current_feb_first_day As Integer: current_feb_first_day = Weekday(current_feb_firstDate)
    Dim next_feb_first_day As Integer: next_feb_first_day = Weekday(next_feb_firstDate)
    Dim may_first_day As Integer: may_first_day = Weekday(may_firstDate)
    Dim aug_first_day As Integer: aug_first_day = Weekday(aug_firstDate)
    Dim nov_first_day As Integer: nov_first_day = Weekday(nov_firstDate)
    
    Dim current_feb_first_week_firstDate As Date: current_feb_first_week_firstDate = Common.FirstDayInWeek(current_feb_firstDate)
    Dim next_feb_first_week_firstDate As Date: next_feb_first_week_firstDate = Common.FirstDayInWeek(next_feb_firstDate)
    Dim may_first_week_firstDate As Date: may_first_week_firstDate = Common.FirstDayInWeek(may_firstDate)
    Dim aug_first_week_firstDate As Date: aug_first_week_firstDate = Common.FirstDayInWeek(aug_firstDate)
    Dim nov_first_week_firstDate As Date: nov_first_week_firstDate = Common.FirstDayInWeek(nov_firstDate)
    
    Dim saturday_day As Integer: saturday_day = 7
    
    Dim current_firstDate_q1 As Date
    Dim next_firstDate_q1 As Date
    Dim lastDate_q1 As Date
    Dim firstDate_q2 As Date
    Dim lastDate_q2 As Date
    Dim firstDate_q3 As Date
    Dim lastDate_q3 As Date
    Dim firstDate_q4 As Date
    Dim lastDate_q4 As Date
    
    If isWalmart Then

        'to get the first date of this year's q1
        current_feb_firstDate = DateValue("2/1/" & current_year)
        next_feb_firstDate = DateValue("2/1/" & current_year + 1)

        current_feb_first_day = Weekday(current_feb_firstDate)
        next_feb_first_day = Weekday(next_feb_firstDate)

        current_feb_first_week_firstDate = Common.FirstDayInWeek(current_feb_firstDate)
        next_feb_first_week_firstDate = Common.FirstDayInWeek(next_feb_firstDate)

        'get 1st date of Wal-Mart Q1 current year
        If current_feb_first_day = saturday_day Then
            current_firstDate_q1 = current_feb_firstDate
        Else
            current_firstDate_q1 = current_feb_first_week_firstDate - 1
        End If

        'if today's date is before the first date of Q1 in this year, then today is still in last year's Q4
        If dateInWeek < current_firstDate_q1 Then current_year = current_year - 1

        current_feb_firstDate = DateValue("2/1/" & current_year)
        next_feb_firstDate = DateValue("2/1/" & current_year + 1)
        may_firstDate = DateValue("5/1/" & current_year)
        aug_firstDate = DateValue("8/1/" & current_year)
        nov_firstDate = DateValue("11/1/" & current_year)

        current_feb_first_day = Weekday(current_feb_firstDate)
        next_feb_first_day = Weekday(next_feb_firstDate)
        may_first_day = Weekday(may_firstDate)
        aug_first_day = Weekday(aug_firstDate)
        nov_first_day = Weekday(nov_firstDate)

        current_feb_first_week_firstDate = Common.FirstDayInWeek(current_feb_firstDate)
        next_feb_first_week_firstDate = Common.FirstDayInWeek(next_feb_firstDate)
        may_first_week_firstDate = Common.FirstDayInWeek(may_firstDate)
        aug_first_week_firstDate = Common.FirstDayInWeek(aug_firstDate)
        nov_first_week_firstDate = Common.FirstDayInWeek(nov_firstDate)


        'get 1st date of Wal-Mart Q1 current year
        If current_feb_first_day = saturday_day Then
            current_firstDate_q1 = current_feb_firstDate
        Else
            current_firstDate_q1 = current_feb_first_week_firstDate - 1
        End If

        'get 1st date of Wal-Mart Q1 next year
        If next_feb_first_day = saturday_day Then
            next_firstDate_q1 = next_feb_firstDate
        Else
            next_firstDate_q1 = next_feb_first_week_firstDate - 1
        End If

        'get 1st date of Wal-Mart Q2
        If may_first_day = saturday_day Then
            firstDate_q2 = may_firstDate
        Else
            firstDate_q2 = may_first_week_firstDate - 1
        End If

        'get 1st date of Wal-Mart Q3
        If aug_first_day = saturday_day Then
            firstDate_q3 = aug_firstDate
        Else
            firstDate_q3 = aug_first_week_firstDate - 1
        End If

        'get 1st date of Wal-Mart Q4
        If nov_first_day = saturday_day Then
            firstDate_q4 = nov_firstDate
        Else
            firstDate_q4 = nov_first_week_firstDate - 1
        End If

        'get the last date of a quarter
        lastDate_q1 = firstDate_q2 - 1
        lastDate_q2 = firstDate_q3 - 1
        lastDate_q3 = firstDate_q4 - 1
        lastDate_q4 = next_firstDate_q1 - 1

        'Wal-Mart Q1
        If dateInWeek >= current_firstDate_q1 And dateInWeek <= lastDate_q1 Then

            getCurrentQuarterNumber = 1
            Exit Function

        'Wal-Mart Q2
        ElseIf dateInWeek >= firstDate_q2 And dateInWeek <= lastDate_q2 Then

            getCurrentQuarterNumber = 2
            Exit Function

        'Wal-Mart Q3
        ElseIf dateInWeek >= firstDate_q3 And dateInWeek <= lastDate_q3 Then

            getCurrentQuarterNumber = 3
            Exit Function

        'Wal-Mart Q4
        ElseIf dateInWeek >= firstDate_q4 And dateInWeek <= lastDate_q4 Then

            getCurrentQuarterNumber = 4
            Exit Function

        End If

    ElseIf Not isWalmart Then
        
        current_firstDate_q1 = DateValue("1/1/" & year(dateInWeek))
        lastDate_q1 = DateValue("4/1/" & year(dateInWeek)) - 1
        firstDate_q2 = DateValue("4/1/" & year(dateInWeek))
        lastDate_q2 = DateValue("7/1/" & year(dateInWeek)) - 1
        firstDate_q3 = DateValue("7/1/" & year(dateInWeek))
        lastDate_q3 = DateValue("10/1/" & year(dateInWeek)) - 1
        firstDate_q4 = DateValue("10/1/" & year(dateInWeek))
        lastDate_q4 = DateValue("12/31/" & year(dateInWeek))
    
        'All other customers Q1
        If dateInWeek >= current_firstDate_q1 And dateInWeek <= lastDate_q1 Then
        
            getCurrentQuarterNumber = 1
            Exit Function
    
        'All other customers Q2
        ElseIf dateInWeek >= firstDate_q2 And dateInWeek <= lastDate_q2 Then
        
            getCurrentQuarterNumber = 2
            Exit Function
            
        'All other customers Q3
        ElseIf dateInWeek >= firstDate_q3 And dateInWeek <= lastDate_q3 Then
        
            getCurrentQuarterNumber = 3
            Exit Function
    
        'All other customers Q4
        ElseIf dateInWeek >= firstDate_q4 And dateInWeek <= lastDate_q4 Then
        
            getCurrentQuarterNumber = 4
            Exit Function
    
        End If
        
        
    End If
    

End Function

'get the date number from the last quarter
Public Function getLastQuarterDate(isWalmart As Boolean, dateInWeek As Date) As Date
    
    Dim current_month As Integer: current_month = month(dateInWeek)
    Dim current_quarter As Integer: current_quarter = getCurrentQuarterNumber(isWalmart, dateInWeek)
    '*******************************************this walmart logic need investigation
    If isWalmart And current_quarter = 4 And current_month <= 2 Then
        getLastQuarterDate = DateValue("10/10/" & year(dateInWeek) - 1)
        Exit Function

    Else
        
        current_month = getLastQuarterMonth(current_month)
        getLastQuarterDate = DateValue(CStr(current_month & "/10/" & year(dateInWeek)))
        Exit Function
    
    End If
    
    
    
End Function


'get the month number of the last quarter by a given month number
Public Function getLastQuarterMonth(current_month As Integer) As Integer
    
    Select Case current_month
        
        Case 1 To 3
            
            getLastQuarterMonth = 9 + current_month
        
        Case 4 To 12
            
            getLastQuarterMonth = current_month - 3
        
    End Select
    

End Function

'get the date number from the last month
Public Function getLastMonthDate(isWalmart As Boolean, dateInWeek As Date) As Date
    
    Dim current_year As Integer: current_year = year(dateInWeek)
    Dim current_month As Integer: current_month = month(dateInWeek)
    Dim last_month As Integer: last_month = getLastMonthNumber(current_month)
    
    getLastMonthDate = DateValue(last_month & "/10/" & current_year)


End Function

'get the month number of the last month
Public Function getLastMonthNumber(current_month As Integer) As Integer

    Select Case current_month
    
        Case 1
            
            getLastMonthNumber = 12
        
        Case Else
        
            getLastMonthNumber = current_month - 1
        
    End Select
    
End Function




