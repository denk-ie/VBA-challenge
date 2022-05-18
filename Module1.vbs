Sub StockOneYear()
   
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
   
        Dim Ticker As String
    
        Dim YearlyChange As Double
        YearlyChange = 0
          
        Dim PercentChange As Double
        
        Dim TotalStockVolume As Double
        TotalStockVolume = 0
        
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        
        Dim YearStart As Long
        YearStart = 2
              
        Dim i As Double
        
        Dim LastRow As Long
        LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                Ticker = ws.Cells(i, 1).Value
                
                If ws.Cells(YearStart, 3).Value = 0 Then
                    
                    For j = YearStart To i
                    
                        If ws.Cells(YearStart, 3).Value <> 0 Then
                        
                            YearStart = j
                    Exit For
                        
                        End If
                    
                    Next j
                
                End If
                
                YearlyChange = ws.Cells(i, 6).Value - ws.Cells(YearStart, 3).Value
                
                If ws.Cells(YearStart, 3).Value = 0 Then
                
                    PercentChange = YearlyChange
                    
                Else
                
                    PercentChange = (YearlyChange / ws.Cells(YearStart, 3).Value)
            
                End If
                
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
                ws.Range("I" & SummaryTableRow).Value = Ticker
                
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                
                ws.Range("K" & SummaryTableRow).Value = PercentChange
                
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                
                ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
                
                SummaryTableRow = SummaryTableRow + 1
                
                YearlyChange = 0
                
                PercentChange = 0
                
                TotalStockVolume = 0
                
                YearStart = i + 1
                
            Else
                
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
            End If
            
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
            
            If ws.Range("J" & i).Value > 0 Then
                
                ws.Range("J" & i).Interior.ColorIndex = 4
                
            ElseIf ws.Range("J" & i).Value < 0 Then
            
                ws.Range("J" & i).Interior.ColorIndex = 3
            
            End If
        Next i
    Next ws
End Sub
    

