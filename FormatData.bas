Attribute VB_Name = "FormatData"
Option Explicit

'OWNER: SICHENG (CHARLES) ZHOU
'CONTACT: CHARLES.ZHOU@DSCLOGISTICS.COM

'add a comma on the right of each value
Public Sub addCommaToRight()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    'insert a column to right
    Call Common.insertColumnToRight(col)
    'set the new column in text format
    Call Common.formatColumnInText(col + 1)
    'copy the header to the new column
    Cells(1, col + 1) = Cells(1, col)
    
    Dim COMMA As String: COMMA = ","
    
    Dim i As Long
    
    'apply formula to every cell paralleling to the current
    For i = 2 To lastRow
        
        If Cells(i, col) = "" Then
            'do nothing
        Else
            Cells(i, col + 1) = CStr(Cells(i, col)) + COMMA
        End If
        
    Next i
    
    Dim rng As Range: Set rng = Range(Cells(2, col + 1), Cells(lastRow, col + 1))
    rng.Select
    
End Sub
