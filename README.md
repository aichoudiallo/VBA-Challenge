# VBA-Challenge
Sub Multiple_year_stock()
For Each ws In Worksheets

ws.Cells(1, 9).Value = "ticker"
ws.Cells(1, 10).Value = "Yearly_change"
ws.Cells(1, 11).Value = "Yearly_percentage"
ws.Cells(1, 12).Value = "Total_Stock_Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, "O").Value = "Greatest_%_Increase"
ws.Cells(3, "O").Value = "Greatest_%_Decrease"
ws.Cells(4, "O").Value = "Greatest_Total_Volume"

Dim ticker As String
Dim Total_Stock_volume As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_percentage As Double
Dim BiggestIncrease As Double
Dim GreatestTotalVolume As Double

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Total_Stock_volume = 0
BiggestIncrease = 0

For i = 2 To 22771
    Total_Stock_volume = Total_Stock_volume + ws.Cells(i, 7).Value
    If ws.Cells(i + 1, 1) = ws.Cells(i, 1) And ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    ticker = ws.Cells(i, 1).Value
    yearly_open = ws.Cells(i, 3).Value
    End If
    
    If ws.Cells(i - 1, 1) = ws.Cells(i, 1) And ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Yearly_close = ws.Cells(i, 6).Value
    'yearly_change = Yearly_close - yearly_open
  
    
    yearly_change = Yearly_close - yearly_open
    
    
    ws.Range("I" & Summary_Table_Row).Value = ticker
    ws.Range("J" & Summary_Table_Row).Value = yearly_change
    If yearly_change > 0 Then
     ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
     Else
     ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
    If yearly_open <> 0 Then
    ws.Range("K" & Summary_Table_Row).Value = FormatPercent((yearly_change / yearly_open), 2)
    Else
    ws.Range("K" & Summary_Table_Row).Value = Null
    End If
    
    If ws.Range("K" & Summary_Table_Row).Value > BiggestIncrease Then
    BiggestIncrease = ws.Range("K" & Summary_Table_Row).Value
    BiggestIncreaseTicker = ws.Range("I" & Summary_Table_Row).Value
    End If
    
    
    If ws.Range("K" & Summary_Table_Row).Value < BiggestDecrease Then
    BiggestDecrease = ws.Range("K" & Summary_Table_Row).Value
    BiggestDecreaseTicker = ws.Range("I" & Summary_Table_Row).Value
    End If
    
    If ws.Range("L" & Summary_Table_Row).Value > GreatestTotalVolume Then
    GreatestTotalVolume = ws.Range("L" & Summary_Table_Row).Value
    GreatestTotalVolumeTicker = ws.Range("I" & Summary_Table_Row).Value
    End If
    
    
    
    ws.Range("L" & Summary_Table_Row).Value = Total_Stock_volume
    
    Summary_Table_Row = Summary_Table_Row + 1
    
    'ticker = 0
    yearly_open = 0
    yearly_change = 0
    Total_Stock_volume = 0
    
    

    
    End If
    Next i
    ws.Cells(2, "P").Value = BiggestIncreaseTicker
    ws.Cells(2, "Q").Value = FormatPercent(BiggestIncrease, 2)
    
    ws.Cells(3, "P").Value = BiggestDecreaseTicker
    ws.Cells(3, "Q").Value = FormatPercent(BiggestDecrease, 2)
    
    ws.Cells(4, "P").Value = GreastestTotalVolumeTicker
    ws.Cells(4, "Q").Value = GreatestTotalVolume
      
      
Next ws
End Sub

