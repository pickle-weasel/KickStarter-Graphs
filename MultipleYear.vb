Sub MultiYear()

'***************************************************************************
'          Moderate
'***************************************************************************

Dim column As Integer
Dim Ticker As String
Dim ChangeYR As Double
Dim ChangePCT As Double
Dim TotalVol As Double
Dim TableRow As Integer
Dim Row As Integer
Dim openRow As Double
Dim open_row As Double

Row = 7
column = 1
TableRow = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
openRow = 2
open_row = 2

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"


For i = 2 To LastRow
    If Cells(i + 1, column).Value <> Cells(i, column).Value Then
        Ticker = Cells(i, column).Value
        open_value = Cells(open_row, column + 2).Value
        
        If open_value = 0 Then
          For m = open_row To i
            open_value = Cells(m, column + 2).Value
            If open_value > 0 Then
                Exit For
            End If
         Next m
            End If
    
        ChangeYR = Cells(i, column + 5).Value - open_value
        ChangePCT = ChangeYR / open_value
        TotalVol = TotalVol + Cells(i, column + 6).Value
        Range("I" & TableRow).Value = Ticker
        Range("J" & TableRow).Value = ChangeYR
        Range("K" & TableRow).Value = ChangePCT
        Range("L" & TableRow).Value = TotalVol
        TableRow = TableRow + 1
        TotalVol = 0
        open_row = (i + 1)
    
    Else
        TotalVol = TotalVol + Cells(i, column + 6).Value
       
    End If
Next i

LastTableRow = Cells(Rows.Count, 11).End(xlUp).Row

For j = 2 To LastTableRow
    If Cells(j, column + 9).Value > 0 Then
        Cells(j, column + 9).Interior.ColorIndex = 4
        
    Else
        Cells(j, column + 9).Interior.ColorIndex = 3
    
    End If
    
Cells(j, column + 10).NumberFormat = "0.00%"

Next j
'***************************************************************************
'          Hard
'***************************************************************************

Dim maxVal As Double
Dim minVal As Double
Dim maxVolume As Double
Dim maxTicker As String
Dim minTicker As String
Dim volumeTicker As String

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

For k = 2 To LastTableRow
    If Cells(k, column + 10).Value > Cells(k + 1, column + 10).Value And _
        (Cells(k, column + 10).Value > maxVal) Then
        maxVal = Cells(k, column + 10).Value
        maxTicker = Cells(k, column + 8).Value
    
    ElseIf (Cells(k, column + 10).Value > Cells(k + 1, column + 10).Value) Then
        maxVal = maxVal
        maxTicker = maxTicker
        
    Else
        maxVal = maxVal
        maxTicker = maxTicker
    
    End If
    
    If Cells(k, column + 10).Value < Cells(k + 1, column + 10).Value And _
        (Cells(k, column + 10).Value < minVal) Then
        minVal = Cells(k, column + 10).Value
        minTicker = Cells(k, column + 8).Value
    
    ElseIf (Cells(k, column + 10).Value < Cells(k + 1, column + 10).Value) Then
        minVal = minVal
        minTicker = minTicker
        
    Else
        minVal = minVal
        minTicker = minTicker
    
    End If
    
    If Cells(k, column + 11).Value > Cells(k + 1, column + 11).Value And _
        (Cells(k, column + 11).Value > maxVolume) Then
        maxVolume = Cells(k, column + 11).Value
        volumeTicker = Cells(k, column + 8).Value
    
    ElseIf (Cells(k, column + 11).Value > Cells(k + 1, column + 11).Value) Then
        maxVolume = maxVolume
        volumeTicker = volumeTicker
        
    Else
        maxVolume = maxVolume
        volumeTicker = volumeTicker
    
    End If

Next k

Range("Q2").Value = maxVal
Range("Q3").Value = minVal
Range("Q4").Value = maxVolume
Range("P2").Value = maxTicker
Range("P3").Value = minTicker
Range("P4").Value = volumeTicker
Range("Q2:Q3").NumberFormat = "0.00%"

End Sub



