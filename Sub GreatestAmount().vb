Sub GreatestAmount()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
            
        MaxPercentIncrease = WorksheetFunction.Max(ws.Range("K2:K290"))
                    
        MaxPercentDecrease = WorksheetFunction.Min(ws.Range("K2:K290"))
                    
        MaxStockVolume = WorksheetFunction.Max(ws.Range("L2:L290"))
                    
        IncreasedTicker = WorksheetFunction.Match(MaxPercentIncrease, ws.Range("K2:K290"), 0)
                    
        DecreasedTicker = WorksheetFunction.Match(MaxPercentDecrease, ws.Range("K2:K290"), 0)
                    
        MaxVolumeTicker = WorksheetFunction.Match(MaxStockVolume, ws.Range("L2:L290"), 0)
                    
        ws.Range("P" & 2).Value = Cells(IncreasedTicker + 1, 9)
                    
        ws.Range("P" & 3).Value = Cells(DecreasedTicker + 1, 9)
                    
        ws.Range("P" & 4).Value = Cells(MaxVolumeTicker + 1, 9)
                    
        ws.Range("Q" & 2).Value = MaxPercentIncrease
                    
        ws.Range("Q" & 3).Value = MaxPercentDecrease
                    
        ws.Range("Q" & 4).Value = MaxStockVolume
                    
        ws.Range("Q" & 2).NumberFormat = "0.00%"
                    
        ws.Range("Q" & 3).NumberFormat = "0.00%"
    Next ws
End Sub
                

