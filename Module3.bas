Attribute VB_Name = "Module3"
Sub StockData():
    
    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim Ticker As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PerChange As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double
        
        Worsheetname = ws.Name
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        Ticker = 2
        
        j = 2
        
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox ("Last row in column A is " & LastRowA)
        
        For i = 2 To LastRowA
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
            
            ws.Cells(Ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                If ws.Cells(Ticker, 10).Value > 0 Then
                
                ws.Cells(Ticker, 10).Interior.ColorIndex = 4
                
                Else
                
                ws.Cells(Ticker, 10).Interior.ColorIndex = 3
                
                End If
                
                If ws.Cells(j, 3).Value <> 0 Then
                
                PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                
                ws.Cells(Ticker, 11).Value = Format(PercentChange, "Percent")
                
                Else
                
                ws.Cells(Ticker, 11).Value = Format(0, "Percent")
                
                End If
            
            ws.Cells(Ticker, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
            Ticker = Ticker + 1
            
            j = i + 1
            
            
            End If
        
        Next i
        
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        GreatVol = ws.Cells(2, 12).Value
        GreatInc = ws.Cells(2, 11).Value
        GreatDec = ws.Cells(2, 11).Value
        
            For i = 2 To LastRowI
                
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                If ws.Cells(i, 11).Value > GreatInc Then
                GreatInc = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatInc = GreatInc
                
                End If
                
                If ws.Cells(i, 11).Value < GreatDec Then
                GreatDec = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDec = GreatDec
                
                End If
                
            ws.Cells(2, 17).Value = Format(GreatInc, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDec, "percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
                
        
    Next ws

End Sub
