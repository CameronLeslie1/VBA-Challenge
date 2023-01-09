# VBA-Challenge

Please see the below VBA Script which calculates the required fields for the Module 2 Challenge.


Sub Module2_Challenge()

' Start Worksheet loop

For Each ws In Worksheets

' Define variables

Dim Ticker As Range
Dim Vol As Range

Set Ticker = Range("A:A")
Set Vol = Range("G:G")

' Count number of days in the year to easily find year end close prices

Dim Days As Integer
Days = Application.WorksheetFunction.CountIf(Range("A:A"), "=AAB")

' Count total rows to analyze for the for statements

  lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
' Add Column and Row Headers

 ws.Cells(1, 9) = "Ticker"
 ws.Cells(1, 10) = "Year Open"
 ws.Cells(1, 11) = "Year Close"
 ws.Cells(1, 12) = "Yearly Change"
 ws.Cells(1, 13) = "Percent Change"
 ws.Cells(1, 14) = "Total Stock Volume"
 ws.Cells(2, 17) = "Greatest % Increase"
 ws.Cells(3, 17) = "Greatest % Decrease"
 ws.Cells(4, 17) = "Greatest Total Volume"
 ws.Cells(1, 18) = "Ticker"
 ws.Cells(1, 19) = "Value"
 
' Loop to find each stock, open price at start of the year and close price at end of the year

  For i = 1 To lastRow
        
        If ws.Cells(i, 2).Value = (ws.Name + "0102") Then
            
            nextemptyrow = ws.Cells(Rows.Count, 9).End(xlUp).Row + 1
            
            ws.Cells(nextemptyrow, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(nextemptyrow, 10).Value = ws.Cells(i, 3).Value
            ws.Cells(nextemptyrow, 11).Value = ws.Cells(i + Days - 1, 6).Value
            nextemptyrow = nextemptyrow + 1
    
        End If

Next i

' Count total number of stocks to analyze change

lastrow1 = Cells(Rows.Count, 9).End(xlUp).Row
 
' loop to find the yearly change, percent change, and total volume

 For j = 2 To lastrow1
   
   ws.Cells(j, 12).Value = ws.Cells(j, 11).Value - ws.Cells(j, 10).Value
   ws.Cells(j, 13).Value = ws.Cells(j, 12).Value / ws.Cells(j, 10).Value
   ws.Cells(j, 14).Value = Application.WorksheetFunction.SumIf(Ticker, ws.Cells(j, 9), Vol)
   
   If ws.Cells(j, 12).Value > 0 Then
        ws.Cells(j, 12).Interior.ColorIndex = 4
        
        Else: ws.Cells(j, 12).Interior.ColorIndex = 3
    End If
             
 Next j
 
' Find greatest increase / decrease in price and greatest total volume
 
 Dim MaxIncrease As Double
 Dim MinIncrease As Double
 Dim MaxVolume As Double

 MaxIncrease = Application.WorksheetFunction.Max(Range("M:M"))
 MinIncrease = Application.WorksheetFunction.Min(Range("M:M"))
 MaxVolume = Application.WorksheetFunction.Max(Range("N:N"))
 
 ws.Cells(2, 19).Value = MaxIncrease
 ws.Cells(3, 19).Value = MinIncrease
 ws.Cells(4, 19).Value = MaxVolume
              
 ' Find tickers for greatest increase / decrease in price and greatest total volume
 
 MaxIncreaseRow = Application.WorksheetFunction.Match(MaxIncrease, Range("M:M"), 0)
 ws.Cells(2, 18) = ws.Cells(MaxIncreaseRow, 9)
 MinIncreaseRow = Application.WorksheetFunction.Match(MinIncrease, Range("M:M"), 0)
 ws.Cells(3, 18) = ws.Cells(MinIncreaseRow, 9)
 MaxVolumeRow = Application.WorksheetFunction.Match(MaxVolume, Range("N:N"), 0)
 ws.Cells(4, 18) = ws.Cells(MaxVolumeRow, 9)
 
 ' Format Cells as %
 Dim PercentChange As Range
 Set PercentChange = ws.Range("M:M")
 PercentChange.NumberFormat = "0.00%"
 Dim MinMaxChange As Range
 Set MinMaxChange = ws.Range("S2:S3")
 MinMaxChange.NumberFormat = "0.00%"
 

Next ws

End Sub
