Attribute VB_Name = "Module1"
Sub stock_calc()

    Dim datacounter As Integer, summaryrow As Integer, beginYearOpen As Double, endyearclose As Double, totalstockvolume As Double, ticker As String
    Dim LastRow As Long, summarycolumn As Integer, LargestIncPercent As Double, LargestDecPercent As Double, LargestStockVol As Double
    Dim LargestIncTicker As String, LargestDecTicker As String, LargestVolTicker As String
    Dim RedColor As Integer, GreenColor As Integer
    
    RedColor = 3
    GreenColor = 10

    
    For Each ws In Worksheets
        
        datacounter = 0
        summaryrow = 2
        summarcolumn = 8
        totalstockvolume = 0
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
        For i = 3 To LastRow
        
            If (ws.Cells(i, 1) = ws.Cells(i - 1, 1)) Then
            
                If (i <> LastRow) Then
                
                    ticker = ws.Cells(i - 1, 1)
                    datacounter = datacounter + 1
                    totalstockvolume = totalstockvolume + ws.Cells(i - 1, 7)
                    If (datacounter = 1) Then
                        beginYearOpen = ws.Cells(i - 1, 3)
                    End If
                
                Else
                    
                    ticker = ws.Cells(i - 1, 1)
                    totalstockvolume = totalstockvolume + ws.Cells(i - 1, 7) + ws.Cells(i, 7)
                    endyearclose = ws.Cells(i, 6)
                
                
                    ws.Cells(summaryrow, 9) = ticker
                    'ws.Cells(summaryrow, 10) = beginYearOpen
                    'ws.Cells(summaryrow, 11) = endyearclose
                    ws.Cells(summaryrow, 10) = endyearclose - beginYearOpen
                    
                    If ((endyearclose - beginYearOpen) < 0) Then
                        
                        ws.Cells(summaryrow, 10).Interior.ColorIndex = RedColor
                        
                    ElseIf ((endyearclose - beginYearOpen) > 0) Then
                    
                        ws.Cells(summaryrow, 10).Interior.ColorIndex = GreenColor
                    
                    End If
                        
                    
                    If (beginYearOpen = 0) Then
                            
                            ws.Cells(summaryrow, 11) = "NA"
                        
                    Else
                            ws.Cells(summaryrow, 11) = Format(((endyearclose - beginYearOpen) / beginYearOpen), "Percent")
                    End If
                    
                    ws.Cells(summaryrow, 12) = totalstockvolume
                
                    
                    datacounter = 0
                    totalstockvolume = 0
                    summaryrow = summaryrow + 1
                
                End If
                
               
             Else
             
                endyearclose = ws.Cells(i - 1, 6)
                totalstockvolume = totalstockvolume + ws.Cells(i - 1, 7)
                
                
                ws.Cells(summaryrow, 9) = ticker
                'ws.Cells(summaryrow, 10) = beginYearOpen
                'ws.Cells(summaryrow, 11) = endyearclose
                ws.Cells(summaryrow, 10) = endyearclose - beginYearOpen
                
                If ((endyearclose - beginYearOpen) < 0) Then
                    
                    ws.Cells(summaryrow, 10).Interior.ColorIndex = RedColor
                    
                ElseIf ((endyearclose - beginYearOpen) > 0) Then
                
                    ws.Cells(summaryrow, 10).Interior.ColorIndex = GreenColor
                
                End If
                    
                
                If (beginYearOpen = 0) Then
                        
                        ws.Cells(summaryrow, 11) = "NA"
                    
                Else
                        ws.Cells(summaryrow, 11) = Format(((endyearclose - beginYearOpen) / beginYearOpen), "Percent")
                End If
                
                ws.Cells(summaryrow, 12) = totalstockvolume
            
                
                datacounter = 0
                totalstockvolume = 0
                summaryrow = summaryrow + 1
                
                           
            End If
            
               
            
        Next i
        
            LargestIncPercent = 0
            LargestDecPercent = 0
            LargestStockVol = 0
            'LargestIncTicker = ""
            'LargestDecTicker = ""
            'LargestVolTicker = ""
            
            ws.Range("O2") = "Greatest % Increase"
            ws.Range("O3") = "Greatest % Decrease"
            ws.Range("O4") = "Greatest Total Volume"
            'ws.Range("P2") = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
            'MsgBox (ws.Name & LargestIncPercent & LargestDecPercent & LargestStockVol)
            
            For j = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
            
                
            
                If (ws.Range("K" & j) <> "NA" And ws.Range("K" & j) > 0 And ws.Range("K" & j) > LargestIncPercent) Then
                
                    LargestIncPercent = ws.Range("K" & j)
                    LargestIncTicker = ws.Range("I" & j)
                    
                End If
                
                If (ws.Range("K" & j) <> "NA" And ws.Range("K" & j) < 0 And ws.Range("K" & j) < LargestDecPercent) Then
                
                    LargestDecTicker = ws.Range("I" & j)
                    LargestDecPercent = ws.Range("K" & j)
                
                End If
                
                If (ws.Range("L" & j) > LargestStockVol) Then
                
                    LargestVolTicker = ws.Range("I" & j)
                    LargestStockVol = ws.Range("L" & j)
                
                End If
            
            Next j
            
            'MsgBox (ws.Name & LargestIncTicker & LargestDecTicker & LargestVolTicker)
            
            ws.Range("P1") = "Ticker"
            ws.Range("P2") = LargestIncTicker
            ws.Range("P3") = LargestDecTicker
            ws.Range("P4") = LargestVolTicker
            
            ws.Range("Q1") = "Value"
            ws.Range("Q2") = Format(LargestIncPercent, "Percent")
            ws.Range("Q3") = Format(LargestDecPercent, "Percent")
            ws.Range("Q4") = LargestStockVol
            
        
            ws.Range("I1") = "Ticker"
            ws.Range("J1") = "Yearly Change"
            ws.Range("K1") = "Percent Change"
            ws.Range("L1") = "Total Stock Volume"
            
            ws.Columns("A:Q").AutoFit
    
    Next ws
    
End Sub


