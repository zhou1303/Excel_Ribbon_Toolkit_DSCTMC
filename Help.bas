Attribute VB_Name = "Help"
Option Explicit


Public Sub ontimeCalculationsRules()
    
    Dim msg As String
    Dim msgBoxTitle As String: msgBoxTitle = "Rules of On Time Calculations"
    Dim getOTFAARule As String
    
    getOTFAARule = "Rule of On Time FAA Calculation:" + vbNewLine
    getOTFAARule = getOTFAARule + vbNewLine
    getOTFAARule = getOTFAARule + "A load or a shipment is on-time to FAA if it is on-time to its earliest delivery appointment. "
    getOTFAARule = getOTFAARule + "Use its Target Delivery (Late) as a delivery appointment if it does not have one."
    getOTFAARule = getOTFAARule + vbNewLine
    
    Dim getOTRADRule As String
    
    getOTRADRule = "Rule of On Time RAD Calculation:" + vbNewLine
    getOTRADRule = getOTRADRule + vbNewLine
    getOTRADRule = getOTRADRule + "A load or a shipment is on-time to RAD if it is on-time to its Target Delivery (Late)."
    getOTRADRule = getOTRADRule + vbNewLine
    
    msg = getOTFAARule + vbNewLine + getOTRADRule + vbNewLine
    
    MsgBox msg, , msgBoxTitle
    
End Sub
