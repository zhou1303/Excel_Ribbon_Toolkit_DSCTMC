Attribute VB_Name = "CostCalculations"
Option Explicit

'OWNER: SICHENG (CHARLES) ZHOU
'CONTACT: CHARLES.ZHOU@DSCLOGISTICS.COM

Global Const CARRIER_TOTAL_CHARGE As String = "Carrier Charge"
Global Const CARRIER_TOTAL_FUEL As String = "Carrier Total Fuel"
Global Const CARRIER_TOTAL_LINE_HAUL As String = "Carrier Total Line Haul"
Global Const CARRIER_TOTAL_OTHER As String = "Carrier Total Other"
Global Const CARRIER_TOTAL_DETENTION As String = "Carrier Total Detention"
Global Const CARRIER_TOTAL_ACCESSORIAL As String = "Carrier Total Accessorial"

Global Const CUSTOMER_TOTAL_CHARGE As String = "Customer Charge"
Global Const CUSTOMER_TOTAL_FUEL As String = "Customer Total Fuel"
Global Const CUSTOMER_TOTAL_LINE_HAUL As String = "Customer Total Line Haul"
Global Const CUSTOMER_TOTAL_OTHER As String = "Customer Total Other"
Global Const CUSTOMER_TOTAL_DETENTION As String = "Customer Total Detention"
Global Const CUSTOMER_TOTAL_ACCESSORIAL As String = "Customer Total Accessorial"

Global Const CARRIER_DISTANCE As String = "Carrier Distance"
Global Const WEIGHT As String = "Weight"
Global Const CUSTOMER_CWT As String = "Customer CWT"
Global Const CARRIER_CWT As String = "Carrier CWT"
Global Const CARRIER_COST_PER_MILE As String = "Carrier $/Mile"
Global Const CUSTOMER_COST_PER_MILE As String = "Customer $/Mile"

Global Const AVG_CUSTOMER_CWT As String = "Avg. Customer CWT"
Global Const AVG_CARRIER_CWT As String = "Avg. Carrier CWT"


'calculate carrier total accessorial charge by adding up carrier total other and carrier total detention

Private Sub calculateCarrierAccessorialSum()
    
    Dim totalOtherCol As Long: totalOtherCol = Common.getColumnNumber(CARRIER_TOTAL_OTHER)
    Dim totalDetentionCol As Long: totalDetentionCol = Common.getColumnNumber(CARRIER_TOTAL_DETENTION)
    
    'do not execute code if any of the required columns cannot be found
    If totalOtherCol * totalDetentionCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = Cells(Rows.Count, totalOtherCol).End(xlUp).row
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol)
        
        lastCol = lastCol + 1
        
        Dim otherRng As Range: Set otherRng = Range(Cells(1, totalOtherCol), Cells(lastRow, totalOtherCol))
        Dim detentionRng As Range: Set detentionRng = Range(Cells(1, totalDetentionCol), Cells(lastRow, totalDetentionCol))
        
        Dim otherArr() As Variant: otherArr = otherRng.Value
        Dim detentionArr() As Variant: detentionArr = detentionRng.Value
        
        If Common.getDimension(otherArr) + Common.getDimension(detentionArr) < 4 Then
            'do nothing
        Else
            
            Dim accessorialArr() As Variant
            ReDim accessorialArr(1 To lastRow, 1 To 1) As Variant
            'copy the header to the first row
            accessorialArr(1, 1) = CARRIER_TOTAL_ACCESSORIAL
            
            Dim thisOther As Double
            Dim thisDetention As Double
            
            Dim i As Long
            
            For i = 2 To lastRow
                
                If otherArr(i, 1) = "" Then
                    thisOther = 0
                Else
                    thisOther = otherArr(i, 1)
                End If
                
                If detentionArr(i, 1) = "" Then
                    thisDetention = 0
                Else
                    thisDetention = detentionArr(i, 1)
                End If
                accessorialArr(i, 1) = thisOther + thisDetention
                
                thisOther = 0
                thisDetention = 0
                
            Next i
            
            Dim accessorialRng As Range: Set accessorialRng = Range(Cells(1, lastCol), Cells(lastRow, lastCol))
            accessorialRng.Value = accessorialArr
            accessorialRng.NumberFormat = "$#,##0.00"
            
        End If
        
    End If
    
End Sub

'calculate customer total accessorial charge by adding up customer total other and customer total detention
Private Sub calculateCustomerAccessorialSum()
    
    Dim totalOtherCol As Long: totalOtherCol = Common.getColumnNumber(CUSTOMER_TOTAL_OTHER)
    Dim totalDetentionCol As Long: totalDetentionCol = Common.getColumnNumber(CUSTOMER_TOTAL_DETENTION)
    
    'do not execute code if any of the required columns cannot be found
    If totalOtherCol * totalDetentionCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = Cells(Rows.Count, totalOtherCol).End(xlUp).row
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol)
        
        lastCol = lastCol + 1
        
        Dim otherRng As Range: Set otherRng = Range(Cells(1, totalOtherCol), Cells(lastRow, totalOtherCol))
        Dim detentionRng As Range: Set detentionRng = Range(Cells(1, totalDetentionCol), Cells(lastRow, totalDetentionCol))
        
        Dim otherArr() As Variant: otherArr = otherRng.Value
        Dim detentionArr() As Variant: detentionArr = detentionRng.Value
        
        If Common.getDimension(otherArr) + Common.getDimension(detentionArr) < 4 Then
            'do nothing
        Else
            
            Dim accessorialArr() As Variant
            ReDim accessorialArr(1 To lastRow, 1 To 1) As Variant
            'copy the header to the first row
            accessorialArr(1, 1) = CUSTOMER_TOTAL_ACCESSORIAL
            
            Dim thisOther As Double
            Dim thisDetention As Double
            
            Dim i As Long
            
            For i = 2 To lastRow
                
                If otherArr(i, 1) = "" Then
                    thisOther = 0
                Else
                    thisOther = otherArr(i, 1)
                End If
                
                If detentionArr(i, 1) = "" Then
                    thisDetention = 0
                Else
                    thisDetention = detentionArr(i, 1)
                End If
                accessorialArr(i, 1) = thisOther + thisDetention
                
                thisOther = 0
                thisDetention = 0
                
            Next i
            
            Dim accessorialRng As Range: Set accessorialRng = Range(Cells(1, lastCol), Cells(lastRow, lastCol))
            accessorialRng.Value = accessorialArr
            accessorialRng.NumberFormat = "$#,##0.00"
            
        End If
        
    End If

End Sub

'calculate carrier total accessorial charge by using carrier charge - carrier total fuel - carrier total line haul

Private Sub calculateCarrierAccessorialDifference()
    
    Dim totalCol As Long: totalCol = Common.getColumnNumber(CARRIER_TOTAL_CHARGE)
    Dim totalFuelCol As Long: totalFuelCol = Common.getColumnNumber(CARRIER_TOTAL_FUEL)
    Dim totalLineHaulCol As Long: totalLineHaulCol = Common.getColumnNumber(CARRIER_TOTAL_LINE_HAUL)
    
    'do not execute code if any of the required columns cannot be found
    If totalCol * totalFuelCol * totalLineHaulCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = Cells(Rows.Count, totalCol).End(xlUp).row
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol)
        
        lastCol = lastCol + 1
        
        Dim totalRng As Range: Set totalRng = Range(Cells(1, totalCol), Cells(lastRow, totalCol))
        Dim fuelRng As Range: Set fuelRng = Range(Cells(1, totalFuelCol), Cells(lastRow, totalFuelCol))
        Dim lineHaulRng As Range: Set lineHaulRng = Range(Cells(1, totalLineHaulCol), Cells(lastRow, totalLineHaulCol))
        
        Dim totalArr() As Variant: totalArr = totalRng.Value
        Dim fuelArr() As Variant: fuelArr = fuelRng.Value
        Dim lineHaulArr() As Variant: lineHaulArr = lineHaulRng.Value
        
        If Common.getDimension(totalArr) + Common.getDimension(fuelArr) _
            + Common.getDimension(lineHaulArr) < 6 Then
            'do nothing
        Else
            
            Dim accessorialArr() As Variant
            ReDim accessorialArr(1 To lastRow, 1 To 1) As Variant
            'copy the header to the first row
            accessorialArr(1, 1) = CARRIER_TOTAL_ACCESSORIAL
            
            Dim thisTotal As Double
            Dim thisFuel As Double
            Dim thisLineHaul As Double
            
            Dim i As Long
            
            For i = 2 To lastRow
                
                If totalArr(i, 1) = "" Then
                    thisTotal = 0
                Else
                    thisTotal = totalArr(i, 1)
                End If
                
                If fuelArr(i, 1) = "" Then
                    thisFuel = 0
                Else
                    thisFuel = fuelArr(i, 1)
                End If
                
                If lineHaulArr(i, 1) = "" Then
                    thisLineHaul = 0
                Else
                    thisLineHaul = lineHaulArr(i, 1)
                End If
                
                accessorialArr(i, 1) = thisTotal - thisFuel - thisLineHaul
                
                thisTotal = 0
                thisFuel = 0
                thisLineHaul = 0
                
            Next i
            
            Dim accessorialRng As Range: Set accessorialRng = Range(Cells(1, lastCol), Cells(lastRow, lastCol))
            accessorialRng.Value = accessorialArr
            accessorialRng.NumberFormat = "$#,##0.00"
            
        End If
        
    End If
    
End Sub

'calculate customer total accessorial charge by using customer charge - customer total fuel - customer total line haul

Private Sub calculateCustomerAccessorialDifference()
    
    Dim totalCol As Long: totalCol = Common.getColumnNumber(CUSTOMER_TOTAL_CHARGE)
    Dim totalFuelCol As Long: totalFuelCol = Common.getColumnNumber(CUSTOMER_TOTAL_FUEL)
    Dim totalLineHaulCol As Long: totalLineHaulCol = Common.getColumnNumber(CUSTOMER_TOTAL_LINE_HAUL)
    
    'do not execute code if any of the required columns cannot be found
    If totalCol * totalFuelCol * totalLineHaulCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = Cells(Rows.Count, totalCol).End(xlUp).row
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol)
        
        lastCol = lastCol + 1
        
        Dim totalRng As Range: Set totalRng = Range(Cells(1, totalCol), Cells(lastRow, totalCol))
        Dim fuelRng As Range: Set fuelRng = Range(Cells(1, totalFuelCol), Cells(lastRow, totalFuelCol))
        Dim lineHaulRng As Range: Set lineHaulRng = Range(Cells(1, totalLineHaulCol), Cells(lastRow, totalLineHaulCol))
        
        Dim totalArr() As Variant: totalArr = totalRng.Value
        Dim fuelArr() As Variant: fuelArr = fuelRng.Value
        Dim lineHaulArr() As Variant: lineHaulArr = lineHaulRng.Value
        
        If Common.getDimension(totalArr) + Common.getDimension(fuelArr) _
            + Common.getDimension(lineHaulArr) < 6 Then
            'do nothing
        Else
            
            Dim accessorialArr() As Variant
            ReDim accessorialArr(1 To lastRow, 1 To 1) As Variant
            'copy the header to the first row
            accessorialArr(1, 1) = CUSTOMER_TOTAL_ACCESSORIAL
            
            Dim thisTotal As Double
            Dim thisFuel As Double
            Dim thisLineHaul As Double
            
            Dim i As Long
            
            For i = 2 To lastRow
                
                If totalArr(i, 1) = "" Then
                    thisTotal = 0
                Else
                    thisTotal = totalArr(i, 1)
                End If
                
                If fuelArr(i, 1) = "" Then
                    thisFuel = 0
                Else
                    thisFuel = fuelArr(i, 1)
                End If
                
                If lineHaulArr(i, 1) = "" Then
                    thisLineHaul = 0
                Else
                    thisLineHaul = lineHaulArr(i, 1)
                End If
                
                accessorialArr(i, 1) = thisTotal - thisFuel - thisLineHaul
                
                thisTotal = 0
                thisFuel = 0
                thisLineHaul = 0
                
            Next i
            
            Dim accessorialRng As Range: Set accessorialRng = Range(Cells(1, lastCol), Cells(lastRow, lastCol))
            accessorialRng.Value = accessorialArr
            accessorialRng.NumberFormat = "$#,##0.00"
            
        End If
        
    End If
    
End Sub

'determine which way to calculate carrier total accessorial
Public Sub calculateCarrierAccessorial()
    
    Dim totalOtherCol As Long: totalOtherCol = Common.getColumnNumber(CARRIER_TOTAL_OTHER)
    Dim totalDetentionCol As Long: totalDetentionCol = Common.getColumnNumber(CARRIER_TOTAL_DETENTION)
    Dim totalCol As Long: totalCol = Common.getColumnNumber(CARRIER_TOTAL_CHARGE)
    Dim totalFuelCol As Long: totalFuelCol = Common.getColumnNumber(CARRIER_TOTAL_FUEL)
    Dim totalLineHaulCol As Long: totalLineHaulCol = Common.getColumnNumber(CARRIER_TOTAL_LINE_HAUL)
    
    If totalOtherCol * totalDetentionCol > 0 Then
        Call calculateCarrierAccessorialSum
    ElseIf totalCol * totalFuelCol * totalLineHaulCol > 0 Then
        Call calculateCarrierAccessorialDifference
    Else
        'do nothing
    End If
    
End Sub
    
'determine which way to calculate customer total accessorial
Public Sub calculateCustomerAccessorial()
    
    Dim totalOtherCol As Long: totalOtherCol = Common.getColumnNumber(CUSTOMER_TOTAL_OTHER)
    Dim totalDetentionCol As Long: totalDetentionCol = Common.getColumnNumber(CUSTOMER_TOTAL_DETENTION)
    Dim totalCol As Long: totalCol = Common.getColumnNumber(CUSTOMER_TOTAL_CHARGE)
    Dim totalFuelCol As Long: totalFuelCol = Common.getColumnNumber(CUSTOMER_TOTAL_FUEL)
    Dim totalLineHaulCol As Long: totalLineHaulCol = Common.getColumnNumber(CUSTOMER_TOTAL_LINE_HAUL)
    
    If totalOtherCol * totalDetentionCol > 0 Then
        Call calculateCustomerAccessorialSum
    ElseIf totalCol * totalFuelCol * totalLineHaulCol > 0 Then
        Call calculateCustomerAccessorialDifference
    Else
        'do nothing
    End If
    
End Sub


'calculate customer cost per hundred weight

Public Sub calculateCustomerCWT()
    
    Dim totalCol As Long: totalCol = Common.getColumnNumber(CUSTOMER_TOTAL_CHARGE)
    Dim weightCol As Long: weightCol = Common.getColumnNumber(WEIGHT)
    
    'do not execute code if any of the required columns cannot be found
    If totalCol * weightCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = Cells(Rows.Count, weightCol).End(xlUp).row
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol)
        
        lastCol = lastCol + 1
        
        Dim totalRng As Range: Set totalRng = Range(Cells(1, totalCol), Cells(lastRow, totalCol))
        Dim weightRng As Range: Set weightRng = Range(Cells(1, weightCol), Cells(lastRow, weightCol))
        
        Dim totalArr() As Variant: totalArr = totalRng.Value
        Dim weightArr() As Variant: weightArr = weightRng.Value
        
        If Common.getDimension(totalArr) + Common.getDimension(weightArr) < 4 Then
            'do nothing
        Else
            
            Dim cwtArr() As Variant
            ReDim cwtArr(1 To lastRow, 1 To 1) As Variant
            'copy the header to the first row
            cwtArr(1, 1) = CUSTOMER_CWT
            
            Dim thisTotal As Double
            Dim thisWeight As Double
            
            Dim i As Long
            
            For i = 2 To lastRow
                
                If totalArr(i, 1) = "" Then
                    thisTotal = 0
                Else
                    thisTotal = totalArr(i, 1)
                End If
                
                If weightArr(i, 1) = "" Then
                    thisWeight = 0
                Else
                    thisWeight = weightArr(i, 1)
                End If
                
                If thisWeight = 0 Then
                    cwtArr(i, 1) = 0
                Else
                    cwtArr(i, 1) = thisTotal / (thisWeight / 100)
                End If
                
                thisTotal = 0
                thisWeight = 0
                
            Next i
            
            Dim cwtRng As Range: Set cwtRng = Range(Cells(1, lastCol), Cells(lastRow, lastCol))
            cwtRng.Value = cwtArr
            cwtRng.NumberFormat = "$#,##0.00"
            
        End If
        
    End If
    
End Sub

'calculate customer average cost per hundread weight
Public Sub calculateAvgCustomerCWT()
    
    Dim totalCol As Long: totalCol = Common.getColumnNumber(CUSTOMER_TOTAL_CHARGE)
    Dim weightCol As Long: weightCol = Common.getColumnNumber(WEIGHT)
    
    'do not execute code if any of the required columns cannot be found
    If totalCol * weightCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = Cells(Rows.Count, weightCol).End(xlUp).row
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol)
        
        lastCol = lastCol + 1
        Cells(1, lastCol) = AVG_CUSTOMER_CWT
        
        Dim totalRng As Range: Set totalRng = Range(Cells(1, totalCol), Cells(lastRow, totalCol))
        Dim weightRng As Range: Set weightRng = Range(Cells(1, weightCol), Cells(lastRow, weightCol))
        
        Dim totalArr() As Variant: totalArr = totalRng.Value
        Dim weightArr() As Variant: weightArr = weightRng.Value
        
        If Common.getDimension(totalArr) + Common.getDimension(weightArr) < 4 Then
            'do nothing
        Else
            
            Dim weightSum, totalSum, avgCWT As Double
            
            Dim thisTotal As Double
            Dim thisWeight As Double
            
            Dim i As Long
            
            For i = 2 To lastRow
                
                If totalArr(i, 1) = "" Then
                    thisTotal = 0
                Else
                    thisTotal = totalArr(i, 1)
                End If
                
                If weightArr(i, 1) = "" Then
                    thisWeight = 0
                Else
                    thisWeight = weightArr(i, 1)
                End If
                
                totalSum = totalSum + thisTotal
                weightSum = weightSum + thisWeight
                
                thisTotal = 0
                thisWeight = 0
                
            Next i
            
            If weightSum = 0 Then
                avgCWT = 0
            Else
                avgCWT = totalSum / (weightSum / 100)
            End If
            
            Dim avgCWTRng As Range: Set avgCWTRng = Range(Cells(2, lastCol), Cells(2, lastCol))
            avgCWTRng.Value = avgCWT
            avgCWTRng.NumberFormat = "$#,##0.00"
            
        End If
        
    End If
    
End Sub

'calculate carrier average cost per hundread weight
Public Sub calculateAvgCarrierCWT()
    
    Dim totalCol As Long: totalCol = Common.getColumnNumber(CARRIER_TOTAL_CHARGE)
    Dim weightCol As Long: weightCol = Common.getColumnNumber(WEIGHT)
    
    'do not execute code if any of the required columns cannot be found
    If totalCol * weightCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = Cells(Rows.Count, weightCol).End(xlUp).row
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol)
        
        lastCol = lastCol + 1
        Cells(1, lastCol) = AVG_CARRIER_CWT
        
        Dim totalRng As Range: Set totalRng = Range(Cells(1, totalCol), Cells(lastRow, totalCol))
        Dim weightRng As Range: Set weightRng = Range(Cells(1, weightCol), Cells(lastRow, weightCol))
        
        Dim totalArr() As Variant: totalArr = totalRng.Value
        Dim weightArr() As Variant: weightArr = weightRng.Value
        
        If Common.getDimension(totalArr) + Common.getDimension(weightArr) < 4 Then
            'do nothing
        Else
            
            Dim weightSum, totalSum, avgCWT As Double
            
            Dim thisTotal As Double
            Dim thisWeight As Double
            
            Dim i As Long
            
            For i = 2 To lastRow
                
                If totalArr(i, 1) = "" Then
                    thisTotal = 0
                Else
                    thisTotal = totalArr(i, 1)
                End If
                
                If weightArr(i, 1) = "" Then
                    thisWeight = 0
                Else
                    thisWeight = weightArr(i, 1)
                End If
                
                totalSum = totalSum + thisTotal
                weightSum = weightSum + thisWeight
                
                thisTotal = 0
                thisWeight = 0
                
            Next i
            
            If weightSum = 0 Then
                avgCWT = 0
            Else
                avgCWT = totalSum / (weightSum / 100)
            End If
            
            Dim avgCWTRng As Range: Set avgCWTRng = Range(Cells(2, lastCol), Cells(2, lastCol))
            avgCWTRng.Value = avgCWT
            avgCWTRng.NumberFormat = "$#,##0.00"
            
        End If
        
    End If
    
End Sub


'calculate carrier cost per hundred weight

Public Sub calculateCarrierCWT()
    
    Dim totalCol As Long: totalCol = Common.getColumnNumber(CARRIER_TOTAL_CHARGE)
    Dim weightCol As Long: weightCol = Common.getColumnNumber(WEIGHT)
    
    'do not execute code if any of the required columns cannot be found
    If totalCol * weightCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = Cells(Rows.Count, weightCol).End(xlUp).row
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol)
        
        lastCol = lastCol + 1
        
        Dim totalRng As Range: Set totalRng = Range(Cells(1, totalCol), Cells(lastRow, totalCol))
        Dim weightRng As Range: Set weightRng = Range(Cells(1, weightCol), Cells(lastRow, weightCol))
        
        Dim totalArr() As Variant: totalArr = totalRng.Value
        Dim weightArr() As Variant: weightArr = weightRng.Value
        
        If Common.getDimension(totalArr) + Common.getDimension(weightArr) < 4 Then
            'do nothing
        Else
            
            Dim cwtArr() As Variant
            ReDim cwtArr(1 To lastRow, 1 To 1) As Variant
            'copy the header to the first row
            cwtArr(1, 1) = CARRIER_CWT
            
            Dim thisTotal As Double
            Dim thisWeight As Double
            
            Dim i As Long
            
            For i = 2 To lastRow
                
                If totalArr(i, 1) = "" Then
                    thisTotal = 0
                Else
                    thisTotal = totalArr(i, 1)
                End If
                
                If weightArr(i, 1) = "" Then
                    thisWeight = 0
                Else
                    thisWeight = weightArr(i, 1)
                End If
                
                If thisWeight = 0 Then
                    cwtArr(i, 1) = 0
                Else
                    cwtArr(i, 1) = thisTotal / (thisWeight / 100)
                End If
                
                thisTotal = 0
                thisWeight = 0
                
            Next i
            
            Dim cwtRng As Range: Set cwtRng = Range(Cells(1, lastCol), Cells(lastRow, lastCol))
            cwtRng.Value = cwtArr
            cwtRng.NumberFormat = "$#,##0.00"
            
        End If
        
    End If
    
End Sub

'calculate customer cost per mile

Public Sub calculateCustomerCostPerMile()
    
    Dim totalCol As Long: totalCol = Common.getColumnNumber(CUSTOMER_TOTAL_CHARGE)
    Dim mileCol As Long: mileCol = Common.getColumnNumber(CARRIER_DISTANCE)
    
    'do not execute code if any of the required columns cannot be found
    If totalCol * mileCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = Cells(Rows.Count, mileCol).End(xlUp).row
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol)
        
        lastCol = lastCol + 1
        
        Dim totalRng As Range: Set totalRng = Range(Cells(1, totalCol), Cells(lastRow, totalCol))
        Dim mileRng As Range: Set mileRng = Range(Cells(1, mileCol), Cells(lastRow, mileCol))
        
        Dim totalArr() As Variant: totalArr = totalRng.Value
        Dim mileArr() As Variant: mileArr = mileRng.Value
        
        If Common.getDimension(totalArr) + Common.getDimension(mileArr) < 4 Then
            'do nothing
        Else
            
            Dim costPerMileArr() As Variant
            ReDim costPerMileArr(1 To lastRow, 1 To 1) As Variant
            'copy the header to the first row
            costPerMileArr(1, 1) = CUSTOMER_COST_PER_MILE
            
            Dim thisTotal As Double
            Dim thisMile As Double
            
            Dim i As Long
            
            For i = 2 To lastRow
                
                If totalArr(i, 1) = "" Then
                    thisTotal = 0
                Else
                    thisTotal = totalArr(i, 1)
                End If
                
                If mileArr(i, 1) = "" Then
                    thisMile = 0
                Else
                    thisMile = mileArr(i, 1)
                End If
                
                If thisMile = 0 Then
                    costPerMileArr(i, 1) = 0
                Else
                    costPerMileArr(i, 1) = thisTotal / thisMile
                End If
                
                thisTotal = 0
                thisMile = 0
                
            Next i
            
            Dim costPerMileRng As Range: Set costPerMileRng = Range(Cells(1, lastCol), Cells(lastRow, lastCol))
            costPerMileRng.Value = costPerMileArr
            costPerMileRng.NumberFormat = "$#,##0.00"
            
        End If
        
    End If
    
End Sub

'calculate carrier cost per mile

Public Sub calculateCarrierCostPerMile()
    
    Dim totalCol As Long: totalCol = Common.getColumnNumber(CARRIER_TOTAL_CHARGE)
    Dim mileCol As Long: mileCol = Common.getColumnNumber(CARRIER_DISTANCE)
    
    'do not execute code if any of the required columns cannot be found
    If totalCol * mileCol = 0 Then
        'do nothing
    Else
        
        Dim lastRow, lastCol As Long
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = Cells(Rows.Count, mileCol).End(xlUp).row
        
        'insert a column to right
        Call Common.insertColumnToRight(lastCol)
        
        lastCol = lastCol + 1
        
        Dim totalRng As Range: Set totalRng = Range(Cells(1, totalCol), Cells(lastRow, totalCol))
        Dim mileRng As Range: Set mileRng = Range(Cells(1, mileCol), Cells(lastRow, mileCol))
        
        Dim totalArr() As Variant: totalArr = totalRng.Value
        Dim mileArr() As Variant: mileArr = mileRng.Value
        
        If Common.getDimension(totalArr) + Common.getDimension(mileArr) < 4 Then
            'do nothing
        Else
            
            Dim costPerMileArr() As Variant
            ReDim costPerMileArr(1 To lastRow, 1 To 1) As Variant
            'copy the header to the first row
            costPerMileArr(1, 1) = CARRIER_COST_PER_MILE
            
            Dim thisTotal As Double
            Dim thisMile As Double
            
            Dim i As Long
            
            For i = 2 To lastRow
                
                If totalArr(i, 1) = "" Then
                    thisTotal = 0
                Else
                    thisTotal = totalArr(i, 1)
                End If
                
                If mileArr(i, 1) = "" Then
                    thisMile = 0
                Else
                    thisMile = mileArr(i, 1)
                End If
                
                If thisMile = 0 Then
                    costPerMileArr(i, 1) = 0
                Else
                    costPerMileArr(i, 1) = thisTotal / thisMile
                End If
                
                thisTotal = 0
                thisMile = 0
                
            Next i
            
            Dim costPerMileRng As Range: Set costPerMileRng = Range(Cells(1, lastCol), Cells(lastRow, lastCol))
            costPerMileRng.Value = costPerMileArr
            costPerMileRng.NumberFormat = "$#,##0.00"
            
        End If
        
    End If
    
End Sub


