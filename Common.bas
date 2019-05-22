Attribute VB_Name = "Common"
Option Explicit
Option Private Module

'OWNER: SICHENG (CHARLES) ZHOU
'CONTACT: CHARLES.ZHOU@DSCLOGISTICS.COM

'get the first row number of a selected range
Public Function getSelectRangeFirstRow() As Long

    getSelectRangeFirstRow = ActiveWindow.RangeSelection.Rows(1).row

End Function

'get the last row number of a selected range
Public Function getSelectRangeLastRow() As Long
    
    Dim firstRow As Long
    firstRow = ActiveWindow.RangeSelection.Rows(1).row
    getSelectRangeLastRow = ActiveWindow.RangeSelection.Rows.Count + firstRow - 1
    
End Function

Public Sub insertColumnToRight(ByVal col As Long)
    
    'move selection to right by one cell
    col = col + 1
    Columns(col).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
End Sub

Public Sub formatColumnInText(ByVal col As Long)
    
    Columns(col).Select
    Selection.NumberFormat = "@"
    
End Sub

'find the column number given a specific title (top row value)
Public Function getColumnNumber(ByVal title As String) As Long
    
    'find the last column of this sheet
    Dim lastCol As Long: lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    Dim col As Long
    'loop column one by one to find the target column with a given title
    For col = 1 To lastCol
        If Cells(1, col) = title Then
            getColumnNumber = col
            Exit For
        Else
            'return value -1 if the column with a given title cannot be found
            getColumnNumber = 0
        End If
    Next col
    
End Function

'get the total number of dimension of an array
Public Function getDimension(ByRef var() As Variant) As Long
    On Error GoTo Err
    Dim i As Long
    Dim tmp As Long
    i = 0
    Do While True
        i = i + 1
        tmp = UBound(var, i)
    Loop
Err:
    getDimension = i - 1
End Function


'get AM/PM of a date variable
Public Function getMidDay(ByVal targetDate As Date) As String
    
    Dim indicator As String: indicator = Right(CStr(TimeValue(targetDate)), 2)
    getMidDay = indicator
    
End Function



'function to return the date of the first day in a given month
Public Function FirstDayInMonth(Optional dtmDate As Variant) As Date
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    FirstDayInMonth = DateSerial( _
        year(dtmDate), month(dtmDate), 1)
End Function

'function to return the date of the last day in a given month
Public Function LastDayInMonth(Optional dtmDate As Variant) As Date
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    LastDayInMonth = DateSerial( _
        year(dtmDate), month(dtmDate) + 1, 0)
End Function

'function to return the date of the first day in a given week
Public Function FirstDayInWeek(Optional dtmDate As Variant) As Date
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    FirstDayInWeek = dtmDate - _
        Weekday(dtmDate, vbUseSystemDayOfWeek) + 1
End Function

'function to return the date of the last day in a given week
Public Function LastDayInWeek(Optional dtmDate As Variant) As Date
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    LastDayInWeek = dtmDate - _
        Weekday(dtmDate, vbUseSystemDayOfWeek) + 7
End Function

'function to return the date of the first day in a given week
Public Function FirstDayInYear(Optional dtmDate As Variant) As Date
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    FirstDayInYear = DateSerial(year(dtmDate), 1, 1)
    
End Function

'function to return the date of the last day in a given year
Public Function LastDayInYear(Optional dtmDate As Variant) As Date
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    LastDayInYear = DateSerial(year(dtmDate), 12, 31)
    
End Function


