Sub multiplestock():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim tickcounter As Long
        Dim lastrowa As Long
        Dim lastrowi As Long
        Dim perchange As Double
        Dim greatinc As Double
        Dim greatdec As Double
        Dim totalvol As Double
        
        WorksheetName = ws.Name

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
       
        tickcounter = 2
        
        j = 2
        
        lastrowa = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To lastrowa
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ws.Cells(tickcounter, 9).Value = ws.Cells(i, 1).Value
                
                ws.Cells(tickcounter, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
            
                    If ws.Cells(tickcounter, 10).Value < 0 Then
                
                    ws.Cells(tickcounter, 10).Interior.ColorIndex = 3
                
                    Else
                
                    ws.Cells(tickcounter, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    If ws.Cells(j, 3).Value <> 0 Then
                    perchange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)

                    ws.Cells(tickcounter, 11).Value = Format(perchange, "Percent")
                    
                    Else
                    
                    ws.Cells(tickcounter, 11).Value = Format(0, "Percent")
                    
                    End If
                    
            
                ws.Cells(tickcounter, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

                tickcounter = tickcounter + 1

                j = i + 1
                
                End If
            
            Next i
            
        lastrowi = ws.Cells(Rows.Count, 9).End(xlUp).Row

        totalvol = ws.Cells(2, 12).Value
        greatinc = ws.Cells(2, 11).Value
        greatdec = ws.Cells(2, 11).Value
        
            
            For i = 2 To lastrowi
            
               
                If ws.Cells(i, 12).Value > totalvol Then
                totalvol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                totalvol = totalvol
                
                End If
                
             
                If ws.Cells(i, 11).Value > greatinc Then
                greatinc = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                greatinc = greatinc
                
                End If
                
                If ws.Cells(i, 11).Value < greatdec Then
                greatdec = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                greatdec = greatdec
                
                End If
                

            ws.Cells(2, 17).Value = Format(greatinc, "Percent")
            ws.Cells(3, 17).Value = Format(greatdec, "Percent")
            ws.Cells(4, 17).Value = Format(totalvol, "Scientific")
            
            Next i
            
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
