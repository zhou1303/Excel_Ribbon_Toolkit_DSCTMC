Attribute VB_Name = "OnTimeCalculations"
Option Explicit

'OWNER: SICHENG (CHARLES) ZHOU
'CONTACT: CHARLES.ZHOU@DSCLOGISTICS.COM

Global Const NA_STRING As String = "#N/A"
Global Const COMMA_STRING As String = ","
Global Const ACTUAL_DELIVERY As String = "Actual Delivery"
Global Const DELIVERY_APPOINTMENT As String = "Delivery Appointment"
Global Const TARGET_DELIVERY_LATE As String = "Target Delivery (Late)"
Global Const CREATE_DATE As String = "Create Date"
Global Const OT_RAD_TITLE As String = "OT RAD(0/1)?"
Global Const OT_FAA_TITLE As String = "OT FAA(0/1)?"

Global Const EMPTY_STRING As String = ""
Global Const SPACE_STRING As String = " "

'get the first date of multiple dates
Private Function getFirstOfMultipleDates(ByVal multiDates As Variant) As Variant
    
    Dim commaPos As Long: commaPos = InStr(1, multiDates, COMMA_STRING, 1)
    getFirstOfMultipleDates = Left(multiDates, commaPos - 1)
    
End Function

'clean dates data for an array
Private Function getFirstOfMultipleDatesValues(ByRef multiDates() As Variant) As Variant()
    
    Dim arrLen As Long: arrLen = UBound(multiDates, 1)
    
    Dim i As Long
    
    For i = 2 To arrLen
        
        If InStr(1, multiDates(i, 1), COMMA_STRING, 1) > 0 Then
            multiDates(i, 1) = getFirstOfMultipleDates(multiDates(i, 1))
        Else
            'do nothing
        End If
        
    Next i
    
    getFirstOfMultipleDatesValues = multiDates
    
End Function

'calculate on time status for given series of times
Private Function getOnTimeRADValues(ByRef actDel() As Variant, _
    ByRef targetDelLate() As Variant) As Variant()
    
    Dim otStatus() As Variant
    
    Dim arrLen As Long: arrLen = UBound(targetDelLate, 1)
    ReDim otStatus(1 To arrLen, 1 To 1) As Variant
    otStatus(1, 1) = OT_RAD_TITLE
    
    Dim i As Long
    
    For i = 2 To arrLen
        
        If actDel(i, 1) = EMPTY_STRING Or targetDelLate(i, 1) = EMPTY_STRING Then
            otStatus(i, 1) = NA_STRING
        Else
        
            If DateValue(actDel(i, 1)) > DateValue(targetDelLate(i, 1)) Then
                otStatus(i, 1) = 0
            Else
                otStatus(i, 1) = 1
            End If
            
        End If
        
    Next i
    
    getOnTimeRADValues = otStatus
    
End Function

'get on time RAD status for objects
Public Sub calculateOnTimeRAD()
    
    Dim actDelCol As Long: actDelCol = Common.getColumnNumber(ACTUAL_DELIVERY)
    Dim targetDelLateCol As Long: targetDelLateCol = Common.getColumnNumber(TARGET_DELIVERY_LATE)
    
    If actDelCol * targetDelLateCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastRow = Cells(Rows.Count, targetDelLateCol).End(xlUp).row
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        
        Dim actDel() As Variant
        Dim targetDelLate() As Variant
        Dim otStatus() As Variant
        
        Dim actDelRng As Range
        Set actDelRng = Range(Cells(1, actDelCol), Cells(lastRow, actDelCol))
        Dim targetDelLateRng As Range
        Set targetDelLateRng = Range(Cells(1, targetDelLateCol), Cells(lastRow, targetDelLateCol))
        
        actDel = actDelRng.Value
        targetDelLate = targetDelLateRng.Value
        
        'stop executing if any range return empty values
        If Common.getDimension(actDel) <= 1 _
            Or Common.getDimension(targetDelLate) <= 1 Then
            'do nothing
        Else
            otStatus = getOnTimeRADValues(actDel, targetDelLate)
        End If
        
    End If
    
    If Common.getDimension(otStatus) <= 1 Then
        'do nothing
    Else
        
        Dim arrLen As Long: arrLen = UBound(otStatus, 1)
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol + 1)
        lastCol = lastCol + 1
        Dim pasteRng As Range: Set pasteRng = Range(Cells(1, lastCol), Cells(arrLen, lastCol))
        pasteRng.Value = otStatus
        
    End If
    
End Sub


'calculate on time status for given series of times
Private Function getOnTimeFAAValues(ByRef actDel() As Variant, _
    ByRef targetDelLate() As Variant, ByRef dlvAppt() As Variant) As Variant()
    
    Dim otStatus() As Variant
    
    'clean delivery appointment values, and get the earliest one
    dlvAppt = getFirstOfMultipleDatesValues(dlvAppt)
    
    Dim arrLen As Long: arrLen = UBound(targetDelLate, 1)
    ReDim otStatus(1 To arrLen, 1 To 1) As Variant
    otStatus(1, 1) = OT_FAA_TITLE
    
    Dim i As Long
    
    For i = 2 To arrLen
        
        If actDel(i, 1) = EMPTY_STRING Or targetDelLate(i, 1) = EMPTY_STRING Then
            otStatus(i, 1) = NA_STRING
        Else
            'if on time to RAD, on time = 1
            If DateValue(actDel(i, 1)) <= DateValue(targetDelLate(i, 1)) Then
                otStatus(i, 1) = 1
            Else
                'if on time to FAA, on time = 1
                If dlvAppt(i, 1) <> EMPTY_STRING Then
                    
                    If DateValue(actDel(i, 1)) <= DateValue(dlvAppt(i, 1)) Then
                        otStatus(i, 1) = 1
                    Else
                        otStatus(i, 1) = 0
                    End If
                    
                Else
                    otStatus(i, 1) = 0
                End If
            End If
            
        End If
        
    Next i
    
    getOnTimeFAAValues = otStatus
    
End Function

'get on time FAA status for objects
Public Sub calculateOnTimeFAA()
    
    Dim actDelCol As Long: actDelCol = Common.getColumnNumber(ACTUAL_DELIVERY)
    Dim targetDelLateCol As Long: targetDelLateCol = Common.getColumnNumber(TARGET_DELIVERY_LATE)
    Dim dlvApptCol As Long: dlvApptCol = Common.getColumnNumber(DELIVERY_APPOINTMENT)
    
    If actDelCol * targetDelLateCol * dlvApptCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastRow = Cells(Rows.Count, targetDelLateCol).End(xlUp).row
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        
        Dim actDel() As Variant
        Dim targetDelLate() As Variant
        Dim dlvAppt() As Variant
        Dim otStatus() As Variant
        
        Dim actDelRng As Range
        Set actDelRng = Range(Cells(1, actDelCol), Cells(lastRow, actDelCol))
        Dim targetDelLateRng As Range
        Set targetDelLateRng = Range(Cells(1, targetDelLateCol), Cells(lastRow, targetDelLateCol))
        Dim dlvApptRng As Range
        Set dlvApptRng = Range(Cells(1, dlvApptCol), Cells(lastRow, dlvApptCol))
        
        actDel = actDelRng.Value
        targetDelLate = targetDelLateRng.Value
        dlvAppt = dlvApptRng.Value
        
        'stop executing if any range return empty values
        If Common.getDimension(actDel) <= 1 _
            Or Common.getDimension(targetDelLate) <= 1 _
            Or Common.getDimension(dlvAppt) <= 1 Then
            'do nothing
        Else
            otStatus = getOnTimeFAAValues(actDel, targetDelLate, dlvAppt)
        End If
        
    End If
    
    If Common.getDimension(otStatus) <= 1 Then
        'do nothing
    Else
        
        Dim arrLen As Long: arrLen = UBound(otStatus, 1)
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol + 1)
        lastCol = lastCol + 1
        Dim pasteRng As Range: Set pasteRng = Range(Cells(1, lastCol), Cells(arrLen, lastCol))
        pasteRng.Value = otStatus
        
    End If
    
End Sub
