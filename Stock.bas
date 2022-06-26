Attribute VB_Name = "Module1"
' account the change of the open price in column C and the close price if Column F
' to get percent change show the change/total


Sub Stock():
For Each ws In Worksheets
    
    Dim WorksheetName As String
    
    'Titles and formating
    ws.Range("I1, P1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("Q1").Value = "Value"
    ws.Range("Q4").NumberFormat = "0"
    ws.Columns("L").NumberFormat = "0"
    ws.Range("P2:Q3").NumberFormat = "0.00%"
    
    'Variables
    'Tickers
    Ticker = ws.Cells(2, 1).Value
    
    'open price
    Op = ws.Cells(2, 3).Value
    
    'closing price
    Cl = 0
    
    'Volumn
    Vol = 0
    
    'Percent
    Per = 0
    
    Dim i As Integer
    'Summary location
    i = 2
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'count to the last row
        For Row = 2 To lastrow
            
            
            If ws.Cells(Row + 1, 1).Value <> Ticker Then
                
                ' Correct order matters
                Cl = WorksheetFunction.Lookup(Ticker, ws.Range("A:A"), ws.Range("F:F"))
                
                'Placement of Ticker and yearly first
                ws.Cells(i, 10).Value = Cl - Op
                ws.Cells(i, 11).Value = (Cl - Op) / Op
                ws.Columns("K").NumberFormat = "0.00%"
                ws.Cells(i, 9).Value = Ticker
                
                
                'Change Ticker
                Ticker = ws.Cells(Row + 1, 1).Value
                            
                ' add the last vol
                Vol = Vol + ws.Cells(Row, 7).Value

                ws.Cells(i, 12).Value = Vol
                
                'reset Vol
                Vol = 0
                
                ' Change Op Value
                Op = ws.Cells(Row + 1, 3).Value
                
                i = i + 1
                
                
            Else
                
                Vol = Vol + ws.Cells(Row, 7).Value
            
            End If
            
        'Max and Min
            
            
        Next Row
    
    'New rows New variables
    colorrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    MaxTotal = WorksheetFunction.Max(ws.Range("L:L"))
    MaxPercent = WorksheetFunction.Max(ws.Range("K:K"))
    MinPercent = WorksheetFunction.min(ws.Range("K:K"))
    
    For Row = 2 To colorrow
    
            'loop for bonus box
            If ws.Cells(Row, 12) = MaxTotal Then
            
                ws.Range("P4") = ws.Cells(Row, 9).Value
            
                ws.Range("Q4") = MaxTotal
            
            ElseIf ws.Cells(Row, 11) = MaxPercent Then
            
                ws.Range("P2") = ws.Cells(Row, 9).Value
            
                ws.Range("Q2") = MaxPercent
            
            ElseIf ws.Cells(Row, 11) = MinPercent Then
            
                ws.Range("P3") = ws.Cells(Row, 9).Value
            
                ws.Range("Q3") = MinPercent

            
            End If
        
            'loop for color
            If ws.Cells(Row, 10).Value > 0# Then
                
                ' adding Color
                ws.Cells(Row, 10).Interior.ColorIndex = 4
                
            ElseIf ws.Cells(Row, 10).Value < 0# Then
                
                ws.Cells(Row, 10).Interior.ColorIndex = 3
                
            
            End If
    
         Next Row
    
    'Autofit all the stuff
    ws.Range("A:Q").Columns.AutoFit
    

    

Next ws

End Sub
