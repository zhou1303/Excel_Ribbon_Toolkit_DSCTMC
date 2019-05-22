Attribute VB_Name = "GeneralTimeParsing"
Option Explicit

'OWNER: SICHENG (CHARLES) ZHOU
'CONTACT: CHARLES.ZHOU@DSCLOGISTICS.COM

Global Const NA_STRING As String = "#N/A"
Global Const COMMA_STRING As String = ","

Global Const EMPTY_STRING As String = ""
Global Const SPACE_STRING As String = " "
Global Const DATE_STRING As String = "Date of"
Global Const TIME_STRING As String = "Time of"
Global Const WEEKDAY_STRING As String = "Wkday of"
Global Const WEEKDAY_NUM_STRING As String = "WkdayNum of"
Global Const MIDDAY_STRING As String = "Midday of"

Public Sub getDate()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'do nothing if there is no data available
    If lastRow = 1 Then Exit Sub
    
    Dim dataArr() As Variant
    Dim rng As Range: Set rng = Range(Cells(1, col), Cells(lastRow, col))
    dataArr = rng.Value
    
    If Common.getDimension(dataArr) <= 1 Then
        'do nothing
    Else
        
        'insert a column to right
        Call Common.insertColumnToRight(col)
        
        Dim arrLen As Long: arrLen = UBound(dataArr, 1)
        Dim dateArr() As Variant
        ReDim dateArr(1 To arrLen, 1 To 1) As Variant
        dateArr(1, 1) = DATE_STRING + SPACE_STRING + dataArr(1, 1)
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = DateValue(dataArr(i, 1))
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
                
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.Value = dateArr
    
End Sub


Public Sub getWeekDay()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'do nothing if there is no data available
    If lastRow = 1 Then Exit Sub
    
    Dim dataArr() As Variant
    Dim rng As Range: Set rng = Range(Cells(1, col), Cells(lastRow, col))
    dataArr = rng.Value
    
    If Common.getDimension(dataArr) <= 1 Then
        'do nothing
    Else
        
        'insert a column to right
        Call Common.insertColumnToRight(col)
        
        Dim arrLen As Long: arrLen = UBound(dataArr, 1)
        Dim dateArr() As Variant
        ReDim dateArr(1 To arrLen, 1 To 1) As Variant
        dateArr(1, 1) = WEEKDAY_STRING + SPACE_STRING + dataArr(1, 1)
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = WeekdayName(Weekday(dataArr(i, 1)))
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
                
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.NumberFormat = "@"
    rng.Value = dateArr

End Sub


Public Sub getWeekDayNumber()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'do nothing if there is no data available
    If lastRow = 1 Then Exit Sub
    
    Dim dataArr() As Variant
    Dim rng As Range: Set rng = Range(Cells(1, col), Cells(lastRow, col))
    dataArr = rng.Value
    
    If Common.getDimension(dataArr) <= 1 Then
        'do nothing
    Else
        
        'insert a column to right
        Call Common.insertColumnToRight(col)
        
        Dim arrLen As Long: arrLen = UBound(dataArr, 1)
        Dim dateArr() As Variant
        ReDim dateArr(1 To arrLen, 1 To 1) As Variant
        dateArr(1, 1) = WEEKDAY_NUM_STRING + SPACE_STRING + dataArr(1, 1)
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = Weekday(dataArr(i, 1))
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
                
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.NumberFormat = "0"
    rng.Value = dateArr
    
End Sub


Public Sub getTime()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'do nothing if there is no data available
    If lastRow = 1 Then Exit Sub
    
    Dim dataArr() As Variant
    Dim rng As Range: Set rng = Range(Cells(1, col), Cells(lastRow, col))
    dataArr = rng.Value
    
    If Common.getDimension(dataArr) <= 1 Then
        'do nothing
    Else
        
        'insert a column to right
        Call Common.insertColumnToRight(col)
        
        Dim arrLen As Long: arrLen = UBound(dataArr, 1)
        Dim dateArr() As Variant
        ReDim dateArr(1 To arrLen, 1 To 1) As Variant
        dateArr(1, 1) = TIME_STRING + SPACE_STRING + dataArr(1, 1)
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = TimeValue(dataArr(i, 1))
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
                
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.NumberFormat = "hh:mm:ss;@"
    rng.Value = dateArr
    
End Sub


Public Sub getMidDay()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'do nothing if there is no data available
    If lastRow = 1 Then Exit Sub
    
    Dim dataArr() As Variant
    Dim rng As Range: Set rng = Range(Cells(1, col), Cells(lastRow, col))
    dataArr = rng.Value
    
    If Common.getDimension(dataArr) <= 1 Then
        'do nothing
    Else
        
        'insert a column to right
        Call Common.insertColumnToRight(col)
        
        Dim arrLen As Long: arrLen = UBound(dataArr, 1)
        Dim dateArr() As Variant
        ReDim dateArr(1 To arrLen, 1 To 1) As Variant
        dateArr(1, 1) = MIDDAY_STRING + SPACE_STRING + dataArr(1, 1)
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = Right(CStr(TimeValue(dataArr(i, 1))), 2)
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
                
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.Value = dateArr
    
End Sub
