Attribute VB_Name = "RemoveAnnotation"
Option Explicit

'OWNER: SICHENG (CHARLES) ZHOU
'CONTACT: CHARLES.ZHOU@DSCLOGISTICS.COM

Global Const DEFAULT_DATA_BEGIN_ROW As Long = 2

'replace a space plus the annotation in parentheses with an empty string.
Private Sub removeAnnotationWithSpace(ByVal col As Long, ByVal annotation As String)
    
    Columns(col).Select
    
    Dim SPACE As String: SPACE = " "
    
    Selection.Replace What:=SPACE + annotation, Replacement:="", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
End Sub

'return as string in a pair of parentheses, including the parentheses.
Private Function getAnnotation(ByVal sample As String) As String
    
    Dim leftParenPos, rightParenPos As Long
    
    leftParenPos = InStr(1, sample, "(", 1)
    rightParenPos = InStr(1, sample, ")", 1)
    
    If leftParenPos = 0 Or rightParenPos = 0 Then
        getAnnotation = ""
    Else
        getAnnotation = Mid(sample, leftParenPos, rightParenPos - leftParenPos + 1)
    End If
    
End Function

'remove annotation in parentheses from a column.
Public Sub removeAnnotationInParentheses()
    
    Dim col As Long
    
    col = ActiveCell.Column
    
    Dim sample As String: sample = Cells(DEFAULT_DATA_BEGIN_ROW, col)
    
    Dim annotation As String: annotation = getAnnotation(sample)
    
    If annotation = "" Then
        'do nothing
    Else
        Call removeAnnotationWithSpace(col, annotation)
    End If
End Sub

'remove annotation in parentheses, and keep column in text format.
Public Sub removeAnnotationInParenthesesAndKeepTextFormat()
    
    Dim col, lastRow As Long
    
    col = ActiveCell.Column
    lastRow = Cells(Rows.Count, col).End(xlUp).row
    
    Dim sample As String: sample = Cells(DEFAULT_DATA_BEGIN_ROW, col)
    
    Dim annotation As String: annotation = getAnnotation(sample)
    
    If annotation = "" Then
        'do nothing
    Else
        'insert a column to right
        Call Common.insertColumnToRight(col)
        'set the new column in text format
        Call Common.formatColumnInText(col + 1)
        'copy the header to the new column
        Cells(1, col + 1) = Cells(1, col)
        
        Dim SPACE As String: SPACE = " "
        
        Dim i As Long
        
        'apply formula to every cell paralleling to the current
        For i = 2 To lastRow
            
            Cells(i, col + 1) = WorksheetFunction.Substitute _
                (Cells(i, col), SPACE + annotation, "")
            
        Next i
    End If
    
End Sub
