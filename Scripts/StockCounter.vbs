Sub QuarterlyChange()

    'declare varaiables
    Dim i, r As Integer
    Dim stockTotal As Double
    Dim LastRow As Long
    Dim openValue, closeValue As Double
    Dim ws As Worksheet
   
    For Each ws In Worksheets
            'column names for aggragated table
            ws.Range("J1").Value = "Ticker"
            ws.Range("K1").Value = "Quarterly Change"
            ws.Range("L1").Value = "Percentage Change"
            ws.Range("M1").Value = "Total Stock Volume"
            
            
        
            'set initial open value
            openValue = ws.Cells(2, 3).Value
            
            'using for row assignment of other table
            r = 2
            
            For i = 2 To ws.Cells(Rows.count, 1).End(xlUp).row
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    'increase stock total by final value
                    stockTotal = stockTotal + ws.Cells(i, 7).Value
                    
                    'capture close value
                    closeValue = ws.Cells(i, 6).Value
        
                    ' move ticker over
                    ws.Cells(r, 10).Value = ws.Cells(i, 1).Value
                    
                    'assign stock total to total stock volume columns
                    ws.Cells(r, 13).Value = stockTotal
                    
                    'show quartlery change
                    ws.Cells(r, 11).Value = closeValue - openValue
                    
                    'get percentage change
                    ws.Cells(r, 12).Value = (closeValue - openValue) / openValue
                    
                    'grab open value for next ticker
                    openValue = ws.Cells(i + 1, 3).Value
             
                    'increase row
                    r = r + 1
                    
                Else:
                    stockTotal = stockTotal + ws.Cells(i, 7).Value
                End If
            
            Next i
            
            'formatting
            ws.Columns("L").NumberFormat = "0.00%"
            ws.Columns("M").NumberFormat = "0"
            
            
            'Calculated Values
            '---------------------------------
            'get last row of new table
            LastRow = ws.Cells(Rows.count, 12).End(xlUp).row
            
            'column names
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Descrease"
            ws.Range("O4").Value = "Greasest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            
            'find highest and lowest values
            Dim high, low As Double
            Dim highTick, lowTick As String
            
            
            'loop through next table
            For i = 2 To LastRow
            
                 'conditonal formating
                If ws.Cells(i, 11).Value > 0 Then
                    ws.Cells(i, 11).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(i, 11).Interior.Color = RGB(255, 0, 0)
                End If
            
            'find greatest % increase, decrease | also find highest total volume
                If ws.Cells(i, 12).Value = ws.Application.WorksheetFunction.Max(Range("L2:L" & LastRow)) Then
                    ws.Range("P2").Value = ws.Cells(i, 10).Value
                    ws.Range("Q2").Value = ws.Cells(i, 12).Value
                ElseIf ws.Cells(i, 12).Value = ws.Application.WorksheetFunction.Min(Range("L2:L" & LastRow)) Then
                    ws.Range("P3").Value = ws.Cells(i, 10).Value
                    ws.Range("Q3").Value = ws.Cells(i, 12).Value
                ElseIf ws.Cells(i, 13).Value = ws.Application.WorksheetFunction.Max(Range("M2:M" & LastRow)) Then
                    ws.Range("P4").Value = ws.Cells(i, 10).Value
                    ws.Range("Q4").Value = ws.Cells(i, 13).Value
                End If
            Next i
            
            'formatting
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            ws.Range("Q4").NumberFormat = "0"
            
            ws.Columns("A:Q").AutoFit
    Next ws

End Sub