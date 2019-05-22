Attribute VB_Name = "TextToColumn"
Option Explicit

'OWNER: SICHENG (CHARLES) ZHOU
'CONTACT: CHARLES.ZHOU@DSCLOGISTICS.COM

Public Sub textToFormat()
    
    Dim row, col As Long
    
    col = ActiveCell.Column
    
    Columns(col).Select
    
    'do nothing if no data is available
    Dim lastRow As Long: lastRow = Cells(Rows.Count, col).End(xlUp).row
    If lastRow = 1 Then Exit Sub
    
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Columns(col), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, COMMA:=False, SPACE:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
End Sub
