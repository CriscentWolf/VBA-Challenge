Sub Stocks()

For Each ws In Worksheets
    
    'set headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'starting variables for each sheet
    Summary_Table_Row = 2
    Ticker = " "
    Open_Price = ws.Cells(2, 3).Value
    Close_Price = 0
    Price_Change = 0
    Price_Change_Percent = 0
    Ticker_volume = 0
    Max_Percent = 0
    Max_Ticker = " "
    Min_Percent = 0
    Min_Ticker = " "
    Max_Volume = 0
    Max_Volume_Ticker = " "
    
        For i = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker = ws.Cells(i, 1).Value
                
                'Calculations
                Close_Price = ws.Cells(i, 6).Value
                Price_Change = Close_Price - Open_Price
                Price_Change_Percent = Price_Change / Open_Price
                Ticker_volume = Ticker_volume + ws.Cells(i, 7).Value
               
               'Outputs
                ws.Cells(Summary_Table_Row, 9).Value = Ticker
                ws.Cells(Summary_Table_Row, 10).Value = Price_Change
                If (Price_Change > 0) Then
                    ws.Cells(Summary_Table_Row, 10).Interior.Color = vbGreen
                ElseIf (Price_Change < 0) Then
                    ws.Cells(Summary_Table_Row, 10).Interior.Color = vbRed
                ElseIf (Price_Change = 0) Then
                    ws.Cells(Summary_Table_Row, 10).Interior.Color = vbYellow
                End If
                ws.Cells(Summary_Table_Row, 11).Value = FormatPercent(Price_Change_Percent)
                ws.Cells(Summary_Table_Row, 12).Value = Ticker_volume
                
                'Next line setup
                Summary_Table_Row = Summary_Table_Row + 1
                Open_Price = ws.Cells(i + 1, 3).Value
                
                'min and max value check
                If (Price_Change_Percent > Max_Percent) Then
                    Max_Percent = Price_Change_Percent
                    Max_Ticker = Ticker
                ElseIf (Price_Change_Percent < Min_Percent) Then
                    Min_Percent = Price_Change_Percent
                    Min_Ticker = Ticker
                End If
                If (Ticker_volume > Max_Volume) Then
                    Max_Volume = Ticker_volume
                    Max_Volume_Ticker = Ticker
                End If
                
                'reset loop
                Price_Change_Percent = 0
                Ticker_volume = 0
            
            'volume totalling
            Else
                Ticker_volume = Ticker_volume + ws.Cells(i, 7).Value
            End If
        Next i
        
    'Final %'s
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("P2").Value = Max_Ticker
    ws.Range("P3").Value = Min_Ticker
    ws.Range("P4").Value = Max_Volume_Ticker
    ws.Range("Q2").Value = FormatPercent(Max_Percent)
    ws.Range("Q3").Value = FormatPercent(Min_Percent)
    ws.Range("Q4").Value = Max_Volume
    
    'Autofit to make Sheets look clean
    ws.Cells.EntireColumn.AutoFit
    ws.Cells.EntireRow.AutoFit
        
Next ws

End Sub

