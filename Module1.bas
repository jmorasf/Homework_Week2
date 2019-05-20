Attribute VB_Name = "Module1"
Sub AllWorksheetsLoop()

Dim ws As Worksheet

' Loop through all of the worksheets in the active workbook.
For Each ws In Worksheets
    ws.Activate
    StockSummary
MsgBox ws.Name
Next

End Sub

Sub StockSummary()
'
' StockSummary Macro
'

Set Sht = ActiveSheet


Dim i As Long
Dim j As Long

Dim currentStock As String
Dim currentTotal As Double
Dim openValue As Double
Dim closeValue As Double
Dim nextStock As String
Dim currentTotalRow As Integer

' Set Headers

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' Values needed before starting the loop

currentTotal = 0
openValue = Cells(2, 3)
j = 2

lastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Loop just like the one done in class - but modified to track the Open/Close stock values

For i = 2 To lastRow
    currentStock = Cells(i, 1).Value
    nextStock = Cells(i + 1, 1).Value
    currentTotal = currentTotal + Cells(i, 7).Value
    
    If currentStock <> nextStock Then
        Cells(j, 9).Value = currentStock
        Cells(j, 12).Value = currentTotal
        Cells(j, 12).NumberFormat = "#,##0"
        closeValue = Cells(i, 6).Value
        Cells(j, 10).Value = closeValue - openValue
        If closeValue > openValue Then
             Cells(j, 10).Interior.Color = vbGreen
        Else
             Cells(j, 10).Interior.Color = vbRed
        End If
        If openValue > 0 Then
            Cells(j, 11).Value = (closeValue - openValue) / openValue
        End If
        Cells(j, 11).NumberFormat = "0.00%"
        
        If i <> lastRow Then
            j = j + 1
            currentTotal = 0
            openValue = Cells(i + 1, 3)
        End If
    
    End If
Next i

' Determine the Maximum/Minimum Values from all stocks

lastRow = Cells(Rows.Count, 9).End(xlUp).Row
Dim minPercentageStock As String
Dim maxPercentageStock As String
Dim maxVolumeStock As String

Dim minPercentage As Double
Dim maxPercentage As Double
Dim maxVolume As Double

' Loop over the previous results

minPercentage = 0
maxPercentage = 0
maxVolume = 0

For i = 2 To lastRow
 
    If Cells(i, 11).Value < minPercentage Then
        minPercentage = Cells(i, 11).Value
        minPercentageStock = Cells(i, 9).Value
    End If
    
    If Cells(i, 11).Value > maxPercentage Then
        maxPercentage = Cells(i, 11).Value
        maxPercentageStock = Cells(i, 9).Value
    End If
    
    If Cells(i, 12).Value > maxVolume Then
        maxVolume = Cells(i, 12).Value
        maxVolumeStock = Cells(i, 9).Value
    End If
 
Next i

' Print /values for last summary

Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"

Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"


Cells(2, 15) = maxPercentageStock
Cells(2, 16) = maxPercentage
Cells(2, 16).NumberFormat = "0.00%"

Cells(3, 15) = minPercentageStock
Cells(3, 16) = minPercentage
Cells(3, 16).NumberFormat = "0.00%"

Cells(4, 15) = maxVolumeStock
Cells(4, 16) = maxVolume
Cells(4, 16).NumberFormat = "#,##0"

Sht.Cells.EntireColumn.AutoFit

End Sub

