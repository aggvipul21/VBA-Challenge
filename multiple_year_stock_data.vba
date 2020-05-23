Attribute VB_Name = "Module1"
Sub stock_calc()

    Dim datacounter As Integer, summaryrow As Integer, beginYearOpen As Double, endyearclose As Double, totalstockvolume As Double, ticker As String
    Dim LastRow As Long, summarycolumn As Integer, LargestIncPercent As Double, LargestDecPercent As Double, LargestStockVol As Double
    Dim LargestIncTicker As String, LargestDecTicker As String, LargestVolTicker As String
    Dim RedColor As Integer, GreenColor As Integer

'Declare variables for color

    RedColor = 3
    GreenColor = 10

'Loop through all the sheets in excel

    For Each ws In Worksheets
        
        datacounter = 0
        summaryrow = 2
        summarcolumn = 8
        totalstockvolume = 0
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
        For i = 3 To LastRow
       
 'Compare ticker value in cell with ticker value in last cell to find change in ticker value in current cell- If values are same this code is run
 
 
            If (ws.Cells(i, 1) = ws.Cells(i - 1, 1)) Then
            
 'Check if this is last row and if not then add stock volume of previous row to stock volume of ticker
 
                If (i <> LastRow) Then
                
                    ticker = ws.Cells(i - 1, 1)
                    datacounter = datacounter + 1
                    totalstockvolume = totalstockvolume + ws.Cells(i - 1, 7)

'If datacounter is 1 it would mean this is the first row for a ticker and assign Open price for Year
                    
                    If (datacounter = 1) Then
                        beginYearOpen = ws.Cells(i - 1, 3)
                    End If
                
'If last row then stock volume of previous row and current row added to get stock volume of ticker and set Year End Price
                
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
                
 'If value of ticker changes from previous row
             Else
            
  'If this is not the first row set end year close price for ticker and add the previous row volume to ticker's stock volume
  
                If (datacounter <> 0) Then
                    endyearclose = ws.Cells(i - 1, 6)
                    totalstockvolume = totalstockvolume + ws.Cells(i - 1, 7)
                
    'If this is the first row set ticker, begin year open price, end year close price for ticker and row volume as ticker's stock volume
    
                Else
                
                    ticker = ws.Cells(i - 1, 1)
                    beginYearOpen = ws.Cells(i - 1, 3)
                    endyearclose = ws.Cells(i - 1, 6)
                    totalstockvolume = totalstockvolume + ws.Cells(i - 1, 7)
                End If

'Setting value of ticker, Yearly change in new cells

                ws.Cells(summaryrow, 9) = ticker
                'ws.Cells(summaryrow, 10) = beginYearOpen
                'ws.Cells(summaryrow, 11) = endyearclose
                ws.Cells(summaryrow, 10) = endyearclose - beginYearOpen
 
 'Setting color for Yearly Change color. If >0 then Green Else If less than 0 then Red else keep default (blank) color
                
                If ((endyearclose - beginYearOpen) < 0) Then
                    
                    ws.Cells(summaryrow, 10).Interior.ColorIndex = RedColor
                    
                ElseIf ((endyearclose - beginYearOpen) > 0) Then
                
                    ws.Cells(summaryrow, 10).Interior.ColorIndex = GreenColor
                
                End If
                    
 'Calculate % change for year's last closing to year's opening price. If year's opening price is 0 then store value as "NA"
                
                If (beginYearOpen = 0) Then
                        
                        ws.Cells(summaryrow, 11) = "NA"
                    
                Else
                        ws.Cells(summaryrow, 11) = Format(((endyearclose - beginYearOpen) / beginYearOpen), "Percent")
                End If
                
 'Assign Total Stock volume
 
                ws.Cells(summaryrow, 12) = totalstockvolume
            
 'Reset datacounter, stockvolumertotal variables to 0 and summary row for data insert for each ticker is incremented by 1
 
                datacounter = 0
                totalstockvolume = 0
                summaryrow = summaryrow + 1
                
                           
            End If
            
               
            
        Next i
        
        
'Calculate Largest % increase, Largest% decrease and greatest total volume for each cell

    'intialize value of variables for Largest%inc, Largest%Dec, TotalStockVol to 0
    
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
            
    'Loop through all records created for ticker, %Change

            For j = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
            
                
      'If %value is not NA and greater than 0 and greater than Largest%inc variable then set Largest%inc variable to new value
      
                If (ws.Range("K" & j) <> "NA" And ws.Range("K" & j) > 0 And ws.Range("K" & j) > LargestIncPercent) Then
                
                    LargestIncPercent = ws.Range("K" & j)
                    LargestIncTicker = ws.Range("I" & j)
                    
                End If
                
       'If %value is not NA and less than 0 and less than Largest%decc variable then set Largest%dec variable to new value
                
                If (ws.Range("K" & j) <> "NA" And ws.Range("K" & j) < 0 And ws.Range("K" & j) < LargestDecPercent) Then
                
                    LargestDecTicker = ws.Range("I" & j)
                    LargestDecPercent = ws.Range("K" & j)
                
                End If
                
        'If value greater than LargestStockVol variable then set LargestStockVol variable to new value
        
                If (ws.Range("L" & j) > LargestStockVol) Then
                
                    LargestVolTicker = ws.Range("I" & j)
                    LargestStockVol = ws.Range("L" & j)
                
                End If
            
            Next j
            
            'MsgBox (ws.Name & LargestIncTicker & LargestDecTicker & LargestVolTicker)
   'Set column names for data inserted through script
            
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


