Attribute VB_Name = "GeneralCalculations"
Option Explicit

'OWNER: SICHENG (CHARLES) ZHOU
'CONTACT: CHARLES.ZHOU@DSCLOGISTICS.COM

Global Const DEST_CITY As String = "Dest City"
Global Const DEST_STATE As String = "Dest State"
Global Const ORIGIN_CITY As String = "Origin City"
Global Const ORIGIN_STATE As String = "Origin State"
Global Const COMMA_STRING As String = ","
Global Const SPACE_STRING As String = " "
Global Const TO_STRING As String = "TO"
Global Const LANE_NAME_STRING As String = "Lane"
Global Const EMPTY_STRING As String = ""

Global Const TARGET_SHIP_LATE As String = "Target Ship (Late)"
Global Const CREATE_DATE As String = "Create Date"
Global Const CARRIER_SHORT_LEAD_TIME_TITLE As String = "Carrier SLT(0/1)?"
Global Const CUSTOMER_SHORT_LEAD_TIME_TITLE As String = "Customer SLT(0/1)?"

Global Const NA_STRING As String = "#N/A"

Global Const SUN_NUM As Integer = 1
Global Const MON_NUM As Integer = 2
Global Const TUE_NUM As Integer = 3
Global Const WED_NUM As Integer = 4
Global Const THU_NUM As Integer = 5
Global Const FRI_NUM As Integer = 6
Global Const SAT_NUM As Integer = 7

Global Const PM_STRING As String = "PM"
Global Const AM_STRING As String = "AM"


Private Function evaluateCarrierShortLeadTimeValue( _
    ByVal targetShipLate As Variant, ByVal createDate As Variant) As Variant
    
    'return #N/A if any of the value is unavailable, or a value contains multiple dates
    If targetShipLate = "" Or createDate = "" _
        Or InStr(1, targetShipLate, COMMA_STRING, 1) > 0 _
        Or InStr(1, createDate, COMMA_STRING, 1) > 0 Then
        evaluateCarrierShortLeadTimeValue = NA_STRING
    Else
        
        Dim createDateDay, dayGap As Integer
        createDateDay = Weekday(createDate)
        
        dayGap = DateValue(targetShipLate) - DateValue(createDate)
        
        Select Case createDateDay
            
            Case SUN_NUM
                If dayGap <= 4 Then evaluateCarrierShortLeadTimeValue = 1
            Case MON_NUM
                If dayGap <= 3 Then evaluateCarrierShortLeadTimeValue = 1
            Case TUE_NUM
                If dayGap <= 3 Then evaluateCarrierShortLeadTimeValue = 1
            Case WED_NUM
                If dayGap <= 5 Then evaluateCarrierShortLeadTimeValue = 1
            Case THU_NUM
                If dayGap <= 5 Then evaluateCarrierShortLeadTimeValue = 1
            Case FRI_NUM
                If dayGap <= 5 Then evaluateCarrierShortLeadTimeValue = 1
            Case SAT_NUM
                If dayGap <= 5 Then evaluateCarrierShortLeadTimeValue = 1
        End Select
        
        If evaluateCarrierShortLeadTimeValue <> 1 _
            And evaluateCarrierShortLeadTimeValue <> NA_STRING Then
            evaluateCarrierShortLeadTimeValue = 0
        End If
        
    End If
    
    
End Function

Private Function getCarrierShortLeadTimeValues( _
    ByRef targetShipLate() As Variant, ByRef createDate() As Variant) As Variant()
    
    Dim arrLen As Long: arrLen = UBound(createDate, 1)
    Dim shortLeadTime() As Variant
    ReDim shortLeadTime(1 To arrLen, 1 To 1) As Variant
    shortLeadTime(1, 1) = CARRIER_SHORT_LEAD_TIME_TITLE
    
    Dim thisCreateDate, thisTargetShipLate As Date
    
    Dim i As Long
    
    For i = 2 To arrLen
        
        thisCreateDate = createDate(i, 1)
        thisTargetShipLate = targetShipLate(i, 1)
        shortLeadTime(i, 1) = evaluateCarrierShortLeadTimeValue(thisTargetShipLate, thisCreateDate)
        
    Next i
    
    getCarrierShortLeadTimeValues = shortLeadTime
    
End Function

Public Sub calculateCarrierShortLeadTime()
    
    Dim createDateCol As Long: createDateCol = Common.getColumnNumber(CREATE_DATE)
    Dim targetShipLateCol As Long: targetShipLateCol = Common.getColumnNumber(TARGET_SHIP_LATE)
    
    
    
    If createDateCol * targetShipLateCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastRow = Cells(Rows.Count, createDateCol).End(xlUp).row
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        
        Dim createDate() As Variant
        Dim targetShipLate() As Variant
        Dim shortLeadTime() As Variant
        
        Dim createDateRng As Range
        Set createDateRng = Range(Cells(1, createDateCol), Cells(lastRow, createDateCol))
        Dim targetShipLateRng As Range
        Set targetShipLateRng = Range(Cells(1, targetShipLateCol), Cells(lastRow, targetShipLateCol))
        
        createDate = createDateRng.Value
        targetShipLate = targetShipLateRng.Value
        
        'stop executing if any range return empty values
        If Common.getDimension(createDate) <= 1 Or Common.getDimension(targetShipLate) <= 1 Then
            'do nothing
        Else
            shortLeadTime = getCarrierShortLeadTimeValues(targetShipLate, createDate)
        End If
        
    End If
    
    If Common.getDimension(shortLeadTime) <= 1 Then
        'do nothing
    Else
        
        Dim arrLen As Long: arrLen = UBound(shortLeadTime, 1)
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol + 1)
        lastCol = lastCol + 1
        Dim pasteRng As Range: Set pasteRng = Range(Cells(1, lastCol), Cells(arrLen, lastCol))
        pasteRng.Value = shortLeadTime
        
    End If
    
End Sub

Private Function evaluateCustomerShortLeadTimeValue( _
    ByVal targetShipLate As Variant, ByVal createDate As Variant) As Variant
    
    'return #N/A if any of the value is unavailable, or a value contains multiple dates
    If targetShipLate = "" Or createDate = "" _
        Or InStr(1, targetShipLate, COMMA_STRING, 1) > 0 Or InStr(1, createDate, COMMA_STRING, 1) > 0 Then
        evaluateCustomerShortLeadTimeValue = NA_STRING
    Else
        
        Dim createDateDay, dayGap As Integer
        Dim createDateMidDay As String
        createDateDay = Weekday(createDate)
        createDateMidDay = Common.getMidDay(createDate)
        
        dayGap = DateValue(targetShipLate) - DateValue(createDate)
        
        Select Case createDateDay
            
            Case SUN_NUM
                If dayGap < 4 Then evaluateCustomerShortLeadTimeValue = 1
            Case MON_NUM
                If createDateMidDay = AM_STRING And dayGap < 3 Then
                    evaluateCustomerShortLeadTimeValue = 1
                ElseIf createDateMidDay = PM_STRING And dayGap < 4 Then
                    evaluateCustomerShortLeadTimeValue = 1
                End If
            Case TUE_NUM
                If createDateMidDay = AM_STRING And dayGap < 3 Then
                    evaluateCustomerShortLeadTimeValue = 1
                ElseIf createDateMidDay = PM_STRING And dayGap < 6 Then
                    evaluateCustomerShortLeadTimeValue = 1
                End If
            Case WED_NUM
                If createDateMidDay = AM_STRING And dayGap < 5 Then
                    evaluateCustomerShortLeadTimeValue = 1
                ElseIf createDateMidDay = PM_STRING And dayGap < 6 Then
                    evaluateCustomerShortLeadTimeValue = 1
                End If
            Case THU_NUM
                If createDateMidDay = AM_STRING And dayGap < 5 Then
                    evaluateCustomerShortLeadTimeValue = 1
                ElseIf createDateMidDay = PM_STRING And dayGap < 6 Then
                    evaluateCustomerShortLeadTimeValue = 1
                End If
            Case FRI_NUM
                If createDateMidDay = AM_STRING And dayGap < 5 Then
                    evaluateCustomerShortLeadTimeValue = 1
                ElseIf createDateMidDay = PM_STRING And dayGap < 6 Then
                    evaluateCustomerShortLeadTimeValue = 1
                End If
            Case SAT_NUM
                If dayGap < 5 Then evaluateCustomerShortLeadTimeValue = 1
        End Select
        
        If evaluateCustomerShortLeadTimeValue <> 1 _
            And evaluateCustomerShortLeadTimeValue <> NA_STRING Then
            evaluateCustomerShortLeadTimeValue = 0
        End If
        
    End If
    
    
End Function

Private Function getCustomerShortLeadTimeValues( _
    ByRef targetShipLate() As Variant, ByRef createDate() As Variant) As Variant()
    
    Dim arrLen As Long: arrLen = UBound(createDate, 1)
    Dim shortLeadTime() As Variant
    ReDim shortLeadTime(1 To arrLen, 1 To 1) As Variant
    shortLeadTime(1, 1) = CUSTOMER_SHORT_LEAD_TIME_TITLE
    
    Dim thisCreateDate, thisTargetShipLate As Date
    
    Dim i As Long
    
    For i = 2 To arrLen
        
        thisCreateDate = createDate(i, 1)
        thisTargetShipLate = targetShipLate(i, 1)
        shortLeadTime(i, 1) = evaluateCustomerShortLeadTimeValue(thisTargetShipLate, thisCreateDate)
        
    Next i
    
    getCustomerShortLeadTimeValues = shortLeadTime
    
    
End Function

Public Sub calculateCustomerShortLeadTime()
    
    Dim createDateCol As Long: createDateCol = Common.getColumnNumber(CREATE_DATE)
    Dim targetShipLateCol As Long: targetShipLateCol = Common.getColumnNumber(TARGET_SHIP_LATE)
    
    If createDateCol * targetShipLateCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastRow = Cells(Rows.Count, createDateCol).End(xlUp).row
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        
        Dim createDate() As Variant
        Dim targetShipLate() As Variant
        Dim shortLeadTime() As Variant
        
        Dim createDateRng As Range
        Set createDateRng = Range(Cells(1, createDateCol), Cells(lastRow, createDateCol))
        Dim targetShipLateRng As Range
        Set targetShipLateRng = Range(Cells(1, targetShipLateCol), Cells(lastRow, targetShipLateCol))
        
        createDate = createDateRng.Value
        targetShipLate = targetShipLateRng.Value
        
        'stop executing if any range return empty values
        If Common.getDimension(createDate) <= 1 Or Common.getDimension(targetShipLate) <= 1 Then
            'do nothing
        Else
            shortLeadTime = getCustomerShortLeadTimeValues(targetShipLate, createDate)
        End If
        
    End If
    
    If Common.getDimension(shortLeadTime) <= 1 Then
        'do nothing
    Else
        
        Dim arrLen As Long: arrLen = UBound(shortLeadTime, 1)
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol + 1)
        lastCol = lastCol + 1
        Dim pasteRng As Range: Set pasteRng = Range(Cells(1, lastCol), Cells(arrLen, lastCol))
        pasteRng.Value = shortLeadTime
        
    End If
    
End Sub

Public Sub concatenateLane()
    
    Dim destCityCol As Long: destCityCol = Common.getColumnNumber(DEST_CITY)
    Dim destStateCol As Long: destStateCol = Common.getColumnNumber(DEST_STATE)
    Dim originCityCol As Long: originCityCol = Common.getColumnNumber(ORIGIN_CITY)
    Dim originStateCol As Long: originStateCol = Common.getColumnNumber(ORIGIN_STATE)
    
    'do not execute code if any of the required columns cannot be found
    If destCityCol * destStateCol * originCityCol * originStateCol = 0 Then
        'do nothing
    Else
    
        Dim lastRow, lastCol As Long
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = Cells(Rows.Count, originCityCol).End(xlUp).row
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol)
        
        lastCol = lastCol + 1
        
        'copy the header to the new column
        Cells(1, lastCol) = LANE_NAME_STRING
        
        Dim i As Long
        Dim lane As String
        
        'concatenate city plus state, and apply lane to the new column
        For i = 2 To lastRow
            
            If Cells(i, originCityCol) <> "" _
                And Cells(i, originStateCol) <> "" _
                And Cells(i, destCityCol) <> "" _
                And Cells(i, destStateCol) <> "" Then
            
                lane = lane + CStr(Cells(i, originCityCol))
                lane = lane + COMMA_STRING
                lane = lane + SPACE_STRING
                lane = lane + CStr(Cells(i, originStateCol))
                lane = lane + SPACE_STRING
                lane = lane + TO_STRING
                lane = lane + SPACE_STRING
                lane = lane + CStr(Cells(i, destCityCol))
                lane = lane + COMMA_STRING
                lane = lane + SPACE_STRING
                lane = lane + CStr(Cells(i, destStateCol))
                
                Cells(i, lastCol) = lane
                
                'clear lane
                lane = EMPTY_STRING
            
            End If
            
        Next i
        
    End If
    
End Sub



