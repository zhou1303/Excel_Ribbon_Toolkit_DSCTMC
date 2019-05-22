Attribute VB_Name = "WalmartTimeParsing"
Option Explicit


'OWNER: SICHENG (CHARLES) ZHOU
'CONTACT: CHARLES.ZHOU@DSCLOGISTICS.COM

Global Const NA_STRING As String = "#N/A"
Global Const COMMA_STRING As String = ","

Global Const EMPTY_STRING As String = ""
Global Const SPACE_STRING As String = " "
Global Const WM_WK_ENDING As String = "Walmart Wk Ending"
Global Const WM_WK_BEGINNING As String = "Walmart Wk Beginning"
Global Const WM_MTH_ENDING As String = "Walmart Mth Ending"
Global Const WM_MTH_BEGINNING As String = "Walmart Mth Beginning"
Global Const WM_QTR_ENDING As String = "Walmart Qtr Ending"
Global Const WM_QTR_BEGINNING As String = "Walmart Qtr Beginning"
Global Const WM_YR_ENDING As String = "Walmart Yr Ending"
Global Const WM_YR_BEGINNING As String = "Walamrt Yr Beginning"
Global Const WM_CUR_MTH As String = "Walmart Cur Mth"
Global Const WM_CUR_QTR As String = "Walmart Cur Qtr"
Global Const WM_CUR_YR As String = "Walmart Cur Yr"

'get the last date of a Walmart week given a certain date
Public Function walmartWeekEnding(ByVal dateInput As Date) As Date
    
    Dim firstDate As Date
    Dim lastDate As Date
    Call WalmartTimeIdentification.getWalmartTimeWeek(dateInput, firstDate, lastDate)
    walmartWeekEnding = lastDate
    
End Function

'get the first date of a Walmart week given a certain date
Public Function walmartWeekBeginning(ByVal dateInput As Date) As Date
    
    Dim firstDate As Date
    Dim lastDate As Date
    Call WalmartTimeIdentification.getWalmartTimeWeek(dateInput, firstDate, lastDate)
    walmartWeekBeginning = firstDate
    
End Function

'get the last date of a Walmart month given a certain date
Public Function walmartMonthEnding(ByVal dateInput As Date) As Date
    
    Dim firstDate As Date
    Dim lastDate As Date
    Call WalmartTimeIdentification.getWalmartTimeMonth(dateInput, firstDate, lastDate)
    walmartMonthEnding = lastDate
    
End Function

'get the last date of a Walmart month given a certain date
Public Function walmartMonthBeginning(ByVal dateInput As Date) As Date
    
    Dim firstDate As Date
    Dim lastDate As Date
    Call WalmartTimeIdentification.getWalmartTimeMonth(dateInput, firstDate, lastDate)
    walmartMonthBeginning = firstDate
    
End Function


'get the last date of a Walmart quarter given a certain date
Public Function walmartQuarterEnding(ByVal dateInput As Date) As Date
    
    Dim firstDate As Date
    Dim lastDate As Date
    Call WalmartTimeIdentification.getWalmartTimeQuarter(dateInput, firstDate, lastDate)
    walmartQuarterEnding = lastDate
    
End Function

'get the last date of a Walmart quarter given a certain date
Public Function walmartQuarterBeginning(ByVal dateInput As Date) As Date
    
    Dim firstDate As Date
    Dim lastDate As Date
    Call WalmartTimeIdentification.getWalmartTimeQuarter(dateInput, firstDate, lastDate)
    walmartQuarterBeginning = firstDate
    
End Function

'get the last date of a Walmart year given a certain date
Public Function walmartYearEnding(ByVal dateInput As Date) As Date
    
    Dim firstDate As Date
    Dim lastDate As Date
    Call WalmartTimeIdentification.getWalmartTimeYear(dateInput, firstDate, lastDate)
    walmartYearEnding = lastDate
    
End Function

'get the last date of a Walmart year given a certain date
Public Function walmartYearBeginning(ByVal dateInput As Date) As Date
    
    Dim firstDate As Date
    Dim lastDate As Date
    Call WalmartTimeIdentification.getWalmartTimeYear(dateInput, firstDate, lastDate)
    walmartYearBeginning = firstDate
    
End Function

'get a group of walmart week endings for a column of dates
Public Sub getWalmartWeekEndings()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'exit sub if pointed column is empty
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
        dateArr(1, 1) = WM_WK_ENDING
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING _
                    And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = walmartWeekEnding(DateValue(dataArr(i, 1)))
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
            
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.NumberFormat = "m/d/yyyy"
    rng.Value = dateArr
    
End Sub


'get a group of walmart week beginnings for a column of dates
Public Sub getWalmartWeekBeginnings()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'exit sub if pointed column is empty
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
        dateArr(1, 1) = WM_WK_BEGINNING
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING _
                    And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = walmartWeekBeginning(DateValue(dataArr(i, 1)))
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
            
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.NumberFormat = "m/d/yyyy"
    rng.Value = dateArr
    
End Sub

'get a group of walmart month beginnings for a column of dates
Public Sub getWalmartMonthEndings()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'exit sub if pointed column is empty
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
        dateArr(1, 1) = WM_MTH_ENDING
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING _
                    And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = walmartMonthEnding(DateValue(dataArr(i, 1)))
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
            
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.NumberFormat = "m/d/yyyy"
    rng.Value = dateArr
    
End Sub


'get a group of walmart month beginnings for a column of dates
Public Sub getWalmartMonthBeginnings()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'exit sub if pointed column is empty
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
        dateArr(1, 1) = WM_MTH_BEGINNING
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING _
                    And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = walmartMonthBeginning(DateValue(dataArr(i, 1)))
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
            
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.NumberFormat = "m/d/yyyy"
    rng.Value = dateArr
    
End Sub


'get a group of walmart quarter beginnings for a column of dates
Public Sub getWalmartQuarterEndings()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'exit sub if pointed column is empty
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
        dateArr(1, 1) = WM_QTR_ENDING
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING _
                    And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = walmartQuarterEnding(DateValue(dataArr(i, 1)))
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
            
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.NumberFormat = "m/d/yyyy"
    rng.Value = dateArr
    
End Sub


'get a group of walmart quarter beginnings for a column of dates
Public Sub getWalmartQuarterBeginnings()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'exit sub if pointed column is empty
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
        dateArr(1, 1) = WM_QTR_BEGINNING
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING _
                    And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = walmartQuarterBeginning(DateValue(dataArr(i, 1)))
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
            
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.NumberFormat = "m/d/yyyy"
    rng.Value = dateArr
    
End Sub

'get a group of walmart year beginnings for a column of dates
Public Sub getWalmartYearEndings()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'exit sub if pointed column is empty
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
        dateArr(1, 1) = WM_YR_ENDING
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING _
                    And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = walmartYearEnding(DateValue(dataArr(i, 1)))
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
            
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.NumberFormat = "m/d/yyyy"
    rng.Value = dateArr
    
End Sub


'get a group of walmart year beginnings for a column of dates
Public Sub getWalmartYearBeginnings()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'exit sub if pointed column is empty
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
        dateArr(1, 1) = WM_YR_BEGINNING
        
        Dim i As Long
        
        For i = 2 To arrLen
            
            If Application.WorksheetFunction.IsNA(dataArr(i, 1)) Then
                dateArr(i, 1) = NA_STRING
            Else
            
                If dataArr(i, 1) <> EMPTY_STRING _
                    And InStr(1, dataArr(i, 1), COMMA_STRING, 1) = 0 Then
                    
                    On Error Resume Next:
                        dateArr(i, 1) = walmartYearBeginning(DateValue(dataArr(i, 1)))
                    
                Else
                
                    dateArr(i, 1) = NA_STRING
                    
                End If
            
            End If
            
        Next i
    
    End If
    
    Set rng = Range(Cells(1, col + 1), Cells(lastRow, col + 1))
    rng.NumberFormat = "m/d/yyyy"
    rng.Value = dateArr
    
End Sub
