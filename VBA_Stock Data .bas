Sub stockdata()
        
    
        For Each ws In Worksheets
        Dim WorksheetName As String
        
        'Present row
        Dim x As Long
        'Ticker block's first row
        Dim y As Long
        'Ticker row index counter
        Dim TickerIndex As Long
        'Column A last row
        Dim LRA As Long
        'Column I last row
        Dim LRI As Long
        'Variable for calculating percent change
        Dim PercentChange As Double
        'Variable for calculating greatest % increase
        Dim GreatPerInc As Double
        'Variable for calculating greatest % decrease
        Dim GreatPerDec As Double
        'Variable for calculating greatest total volume
        Dim GTV As Double
        
       'For WorksheetName
        WorksheetName = ws.Name
        
        'Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Set the first row of the Ticker Counter
        TickerIndex = 2
        
        'Set starting row to 2
        y = 2
        
        'To Find the last filled cell in column A
        LRA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
                   'For looping
                   For x = 2 To LRA
            
                  'It Checks when ticker name changes
                   If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
                
                   'When ticker name changes it write ticker in column I
                    ws.Cells(TickerIndex, 9).Value = ws.Cells(x, 1).Value
                 
                    'Calculate and write the values in column J
                    ws.Cells(TickerIndex, 10).Value = ws.Cells(x, 6).Value - ws.Cells(y, 3).Value
                
                    'Conditional formating
                    If ws.Cells(TickerIndex, 10).Value < 0 Then
                
                    'Highlights -Ve change in Red
                    ws.Cells(TickerIndex, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Highlights +Ve change in green
                    ws.Cells(TickerIndex, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write percentage value in column K
                    If ws.Cells(y, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(x, 6).Value - ws.Cells(y, 3).Value) / ws.Cells(y, 3).Value)
                    
                    'Changing in Percent
                    ws.Cells(TickerIndex, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickerIndex, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate and write value in column L
                ws.Cells(TickerIndex, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(y, 7), ws.Cells(x, 7)))
                
                'Increase TickerIndex by 1
                TickerIndex = TickerIndex + 1
                
                y = x + 1
                
                End If
            
                Next x
            
            
         'To Find the last filled cell in column I
        
          LRI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        
       
        
                'For looping
                 For x = 2 To LRI
            
                'For greatest total volume, it will check if the value is larg or not and if its larger than it will take a new value
                If ws.Cells(x, 12).Value > GTV Then
                GTV = ws.Cells(x, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(x, 9).Value
                
                Else
                
                GTV = GTV
                
                End If
                
                'For greatest percent increase, it will check if the value is larg or not and if its larger than it will take a new value
                If ws.Cells(x, 11).Value > GreatPerInc Then
                GreatPerInc = ws.Cells(x, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(x, 9).Value
                
                Else
                
                GreatPerInc = GreatPerInc
                
                End If
                
                'For greatest percent decrease, it will check if the value is small or not and if its smaller than it will take a new value
                If ws.Cells(x, 11).Value < GreatPerDec Then
                GreatPerDec = ws.Cells(x, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(x, 9).Value
                
                Else
                
                GreatPerDec = GreatPerDec
                
                End If
                
            'For final results
            ws.Range("Q2").Value = Format(GreatPerInc, "Percent")
            ws.Range("Q3").Value = Format(GreatPerDec, "Percent")
            ws.Range("Q4").Value = Format(GTV, "Scientific")
            
            Next x
            Next ws
End Sub
