Sub Stock_Data_Analysis()

    'Loop through each worksheet
    For Each ws In Worksheets
    
        'Deffining variables
        
        Dim WSname As String
        Dim i As Long
        Dim j As Long
        Dim TickerRowCount As Long
        Dim ColALastRow As Long
        Dim ColILastRow As Long
        Dim PercentChangeCal As Double
        Dim IncreaseCal As Double
        Dim DecreaseCal As Double
        Dim TotalVolume As Double
        
        
        WSname = ws.Name
        
        'Coloumn headers
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
		TickerRowCount = 2
        
        j = 2
        
        'Getting lastrow of Coloumn A
        ColALastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop to summerise ticker
        
        For i = 2 To ColALastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickerRowCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TickerRowCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                
                'Conditional Formating
                
                If ws.Cells(TickerRowCount, 10).Value < 0 Then
                    ws.Cells(TickerRowCount, 10).Interior.ColorIndex = 3
                
                Else
                    ws.Cells(TickerRowCount, 10).Interior.ColorIndex = 4
                    
                End If
                
                
                If ws.Cells(j, 3).Value <> 0 Then
                    PercentChangeCal = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3))
                    ws.Cells(TickerRowCount, 11).Value = Format(PercentChangeCal, "Percent")
                Else
                    ws.Cells(TickerRowCount, 11).Value = Format(0, "Percent")
                    
                End If
                
                
                ws.Cells(TickerRowCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                
                TickerRowCount = TickerRowCount + 1
                
                j = i + 1
                
            End If
            
        Next i
        
            'Hard solution
                
           ColILastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
           
           
            IncreaseCal = ws.Cells(2, 11).Value
                          
            DecreaseCal = ws.Cells(2, 11).Value
            
            TotalVolume = ws.Cells(2, 11).Value
            
            
            'Calculate Total volume
            
            For i = 2 To ColILastRow
                If ws.Cells(i, 12).Value > TotalVolume Then
                    TotalVolume = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                    
                Else
                    TotalVolume = TotalVolume
            
                End If
                
                If ws.Cells(i, 11).Value > IncreaseCal Then
                    IncreaseCal = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    
                Else
                    IncreaseCal = IncreaseCal
            
                End If
                
                 If ws.Cells(i, 11).Value < DecreaseCal Then
                    DecreaseCal = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    
                Else
                    DecreaseCal = DecreaseCal
            
                End If
                
                'Display Final Values
                
                ws.Cells(2, 17).Value = Format(IncreaseCal, "Percent")
                ws.Cells(3, 17).Value = Format(DecreaseCal, "Percent")
                ws.Cells(4, 17).Value = Format(TotalVolume, "Scientific")
                
            Next i
            
            Worksheets(WSname).Columns("A:Z").AutoFit
            
        Next ws
            
        
End Sub



