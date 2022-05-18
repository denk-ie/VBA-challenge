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
                
                If ws.Range("J" & SummaryTableRow).Value > 0 Then
    
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                
                ElseIf ws.Range("J" & SummaryTableRow).Value < 0 Then
            
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
    
                Else
                    
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 0

End If
                
                SummaryTableRow = SummaryTableRow + 1
                
                YearlyChange = 0
                
                PercentChange = 0
                
                TotalStockVolume = 0
                
                YearStart = i + 1
                
            Else
                
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
            End If
        Next i
    Next ws
End Sub

