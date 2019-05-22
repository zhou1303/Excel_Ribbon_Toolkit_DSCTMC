Attribute VB_Name = "LeadingZeros"
Option Explicit

'OWNER: SICHENG (CHARLES) ZHOU
'CONTACT: CHARLES.ZHOU@DSCLOGISTICS.COM

Global Const DEFAULT_DATA_BEGIN_ROW As Long = 2
Global Const EMPTY_STRING As String = ""
Global Const ONE_LEADING_ZERO As String = "0"
Global Const TWO_LEADING_ZEROS As String = "00"
Global Const FOUR_LEADING_ZEROS As String = "0000"

'add one leading zero to selected cells.
Public Sub addSelectOneLeadingZero()
    
    Dim col, lastRow As Long
    Dim rngFirstRow, rngLastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    rngFirstRow = Common.getSelectRangeFirstRow
    rngLastRow = Common.getSelectRangeLastRow
    
    'set the new column in text format
    Call Common.formatColumnInText(col)
    
    Dim i, j As Long
    
    'add zero(s) to cells in selected range
    For i = rngFirstRow To lastRow
        
        If i >= rngFirstRow And i <= rngLastRow Then
            Cells(i, col) = ONE_LEADING_ZERO & CStr(Cells(i, col))
        End If
        
    Next i
    
    Dim rng As Range: Set rng = Range(Cells(rngFirstRow, col), Cells(rngLastRow, col))
    rng.Select
    
End Sub


'add two leading zeros to selected cells.
Public Sub addSelectTwoLeadingZeros()
    
    Dim col, lastRow As Long
    Dim rngFirstRow, rngLastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    rngFirstRow = Common.getSelectRangeFirstRow
    rngLastRow = Common.getSelectRangeLastRow
    
    'set the new column in text format
    Call Common.formatColumnInText(col)
    
    Dim i, j As Long
    
    'add zero(s) to cells in selected range
    For i = rngFirstRow To lastRow
        
        If i >= rngFirstRow And i <= rngLastRow Then
            Cells(i, col) = TWO_LEADING_ZEROS & CStr(Cells(i, col))
        End If
        
    Next i
    
    Dim rng As Range: Set rng = Range(Cells(rngFirstRow, col), Cells(rngLastRow, col))
    rng.Select
    
End Sub


'add four leading zeros to selected cells.
Public Sub addSelectFourLeadingZeros()
    
    Dim col, lastRow As Long
    Dim rngFirstRow, rngLastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    rngFirstRow = Common.getSelectRangeFirstRow
    rngLastRow = Common.getSelectRangeLastRow
    
    'set the new column in text format
    Call Common.formatColumnInText(col)
    
    Dim i, j As Long
    
    'add zero(s) to cells in selected range
    For i = rngFirstRow To lastRow
        
        If i >= rngFirstRow And i <= rngLastRow Then
            Cells(i, col) = FOUR_LEADING_ZEROS & CStr(Cells(i, col))
        End If
    Next i
    
    Dim rng As Range: Set rng = Range(Cells(rngFirstRow, col), Cells(rngLastRow, col))
    rng.Select
    
End Sub

'add one leading zero to selected cells.
Public Sub rmSelectLeadingZeros()
    
    Dim col, lastRow As Long
    Dim rngFirstRow, rngLastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    rngFirstRow = Common.getSelectRangeFirstRow
    rngLastRow = Common.getSelectRangeLastRow
    
    'set the new column in text format
    Call Common.formatColumnInText(col)
    
    Dim i, j As Long
    Dim buffString As String
    Dim ifLeadingZero As Boolean
    
    'remove zero(s) to cells in selected range
    For i = rngFirstRow To lastRow
        
        If i <= rngLastRow Then
        
            buffString = CStr(Cells(i, col))
            ifLeadingZero = ifFirstLeadingZero(buffString)
            
            Do While ifLeadingZero
                
                buffString = rmFirstChar(buffString)
                ifLeadingZero = ifFirstLeadingZero(buffString)
                
            Loop
            
            Cells(i, col) = buffString
        
        End If
        
    Next i
    
    Dim rng As Range: Set rng = Range(Cells(rngFirstRow, col), Cells(rngLastRow, col))
    rng.Select
    
End Sub

'check if the first character of a string is a zero
Private Function ifFirstLeadingZero(ByVal original As String) As Boolean
    
    If original = EMPTY_STRING Then
        ifFirstLeadingZero = False
    Else
        
        If Left(original, 1) = ONE_LEADING_ZERO Then
            ifFirstLeadingZero = True
        Else
            ifFirstLeadingZero = False
        End If
    End If
    
End Function

'remove the first character from a string
Private Function rmFirstChar(ByVal original As String) As String
    
    rmFirstChar = Right(original, Len(original) - 1)
    
End Function

