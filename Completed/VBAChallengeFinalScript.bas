Attribute VB_Name = "Module1"
Sub VBAChallengeFinal()

'Declare Variables

Dim ws As Worksheet
Dim Ticker As String
Dim TotalTickers As Integer
Dim LastRow As Long
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim YearlyChange As Double
Dim PercentChange As String
Dim TotalStockVolume As Double
Dim GreatestPercentIncrease As Double
Dim GreatestPercentIncreaseTicker As String
Dim GreatestPercentDecrease As Double
Dim GreatestPercentDecreaseTicker As String
Dim GreatestStockVolume As Double
Dim GreatestStockVolumeTicker As String

'Loop through all the stocks for one year

For Each ws In ActiveWorkbook.Worksheets

    ' Find the last row for each worksheet

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Add headers for Ticker, Yearly Change, Percent Change, and Total Stock Volume to each worksheet. Also add headers for Bonus section.

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ' Set up Variables
    
    Ticker = ""
    TotalTickers = 0
    OpeningPrice = 0
    ClosingPrice = 0
    YearlyChange = 0
    PercentChange = 0
    TotalStockVolume = 0
    
    ' Start looping through each ticker to obtain values for Ticker, yearly change, percent change and total stock volume
    
    For i = 2 To LastRow
    
        Ticker = ws.Cells(i, 1).Value
        
        If OpeningPrice = 0 Then
            
            OpeningPrice = ws.Cells(i, 3).Value
        
        End If
        
        'Calculate Total Stock Volume
        
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
        ' Separate unique Tickers and calculate Yearly Change and percent change apply conditional formatting.
        
        If ws.Cells(i + 1, 1) <> Ticker Then
        
            TotalTickers = TotalTickers + 1
            ws.Cells(TotalTickers + 1, 9) = Ticker
            
            ClosingPrice = ws.Cells(i, 6)
            
            'Yearly Change Calculation
            
            YearlyChange = ClosingPrice - OpeningPrice
            
            ws.Cells(TotalTickers + 1, 10).Value = YearlyChange
            
            'Conditional Formatting for Yearly Change based on value.
            
            If YearlyChange > 0 Then
                
                ws.Cells(TotalTickers + 1, 10).Interior.ColorIndex = 4
                
            ElseIf YearlyChange < 0 Then
            
                ws.Cells(TotalTickers + 1, 10).Interior.ColorIndex = 3
                
            Else
            
                ws.Cells(TotalTickers + 1, 10).Interior.ColorIndex = 6
                
            End If
            
            ' Percent Change Calculation
            
            If OpeningPrice = 0 Then
                
                PercentChange = 0
            
            Else
            
                PercentChange = (YearlyChange / OpeningPrice) * 100 & "%"
                
            End If
            
            ' Reset opening price
            
            OpeningPrice = 0
            
            ws.Cells(TotalTickers + 1, 11).Value = PercentChange
            
            ' Add total Stock volume
            ws.Cells(TotalTickers + 1, 12).Value = TotalStockVolume
            
            'Reset Stock Volume
            
            TotalStockVolume = 0
            
            
        End If
            
    Next i
    
    'Bonus section to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
    
    GreatestPercentIncrease = ws.Cells(2, 11).Value
    GreatestPercentIncreaseTicker = ws.Cells(2, 9).Value
    GreatestPercentDecrease = ws.Cells(2, 11).Value
    GreatestPercentDecreaseTicker = ws.Cells(2, 9).Value
    GreatestTotalVolume = ws.Cells(2, 12).Value
    GreatestTotalVolumeTicker = ws.Cells(2, 9).Value
    
    'Loop through each ticker.
    For i = 2 To LastRow
    
        ' Greatest percent increase.
        
        If ws.Cells(i, 11).Value > GreatestPercentIncrease Then
            
            GreatestPercentIncrease = ws.Cells(i, 11).Value
            
            GreatestPercentIncreaseTicker = ws.Cells(i, 9).Value
        
        End If
        
        ' Greatest percent decrease.
        
        If ws.Cells(i, 11).Value < GreatestPercentDecrease Then
            
            GreatestPercentDecrease = ws.Cells(i, 11).Value
            
            GreatestPercentDecreaseTicker = ws.Cells(i, 9).Value
        
        End If
        
        ' Greatest total volume.
        
        If ws.Cells(i, 12).Value > GreatestTotalVolume Then
            
            GreatestTotalVolume = ws.Cells(i, 12).Value
            
            GreatestTotalVolumeTicker = ws.Cells(i, 9).Value
        
        End If
        
    Next i
    
    ' Add bonus values to each sheet.
    
    ws.Cells(2, 16).Value = GreatestPercentIncreaseTicker
    ws.Cells(3, 16).Value = GreatestPercentDecreaseTicker
    ws.Cells(4, 16).Value = GreatestTotalVolumeTicker
    ws.Cells(2, 17).Value = Format(GreatestPercentIncrease, "Percent")
    ws.Cells(3, 17).Value = Format(GreatestPercentDecrease, "Percent")
    ws.Cells(4, 17).Value = GreatestTotalVolume
    
    
Next ws


End Sub
