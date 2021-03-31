Attribute VB_Name = "Module1"
Sub Stock_Year():

    Dim ws As Worksheet
    For Each ws In Worksheets
    
    Dim SummaryRow As Double
    SummaryRow = 2
    Dim LastRow As Double
    Dim Percent_Change As Double

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
       
j = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
            
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ws.Cells(SummaryRow, 9).Value = ws.Cells(i, 1).Value
                
                ws.Cells(SummaryRow, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                If ws.Cells(SummaryRow, 10).Value < 0 Then
                
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                
                Else
                
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                
                End If
                    
                If ws.Cells(j, 3).Value <> 0 Then
                Percent_Change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                ws.Cells(SummaryRow, 11).Value = Format(Percent_Change, "Percent")
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
                    
                End If
                    
                ws.Cells(SummaryRow, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                SummaryRow = SummaryRow + 1
                
                j = i + 1
                
                End If
            
            Next i

         ws.Range("A:L").Columns.AutoFit
            
    Next ws
    
End Sub

