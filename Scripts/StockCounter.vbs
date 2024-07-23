Sub QuarterlyChange()

    'declare variables
    Dim i, r As Integer
    Dim stockTotal As Double
    Dim LastRow As Long
    Dim openValue, closeValue As Double
    
    'column names for aggragated table
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Quarterly Change"
    Range("L1").Value = "Percentage Change"
    Range("M1").Value = "Total Stock Volume"
    
    
    'set initial open value
    openValue = Cells(2, 3).Value
    
    'using for row assignment of other table
    r = 2
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).row
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'increase stock total by final value
            stockTotal = stockTotal + Cells(i, 7).Value
            
            'capture close value
            closeValue = Cells(i, 6).Value

            ' move ticker over
            Cells(r, 10).Value = Cells(i, 1).Value
            
            'assign stock total to total stock volume columns
            Cells(r, 13).Value = stockTotal
            
            'show quartlery change
            Cells(r, 11).Value = closeValue - openValue
            
            'get percentage change
            Cells(r, 12).Value = (closeValue - openValue) / openValue
            
            'grab open value for next ticker
            openValue = Cells(i + 1, 3).Value
     
            'increase row
            r = r + 1
            
        Else:
            stockTotal = stockTotal + Cells(i, 7).Value
        End If
    
    Next i
    
    'formatting
    Columns("L").NumberFormat = "0.00%"
    Columns("M").NumberFormat = "0"
    
    
    'Calculated Values
    '---------------------------------
    'get last row of new table
    LastRow = Cells(Rows.Count, 12).End(xlUp).row
    
    'column names
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Descrease"
    Range("O4").Value = "Greasest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'find highest and lowest values
    Dim high, low As Double
    Dim highTick, lowTick As String
    
    
    'loop through next table
    For i = 2 To LastRow
    
         'conditonal formating
        If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.Color = RGB(0, 255, 0)
        Else
            Cells(i, 11).Interior.Color = RGB(255, 0, 0)
        End If
    
    'find greatest % increase, decrease | also find highest total volume
        If Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & LastRow)) Then
            Range("P2").Value = Cells(i, 10).Value
            Range("Q2").Value = Cells(i, 12).Value
        ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Min(Range("L2:L" & LastRow)) Then
            Range("P3").Value = Cells(i, 10).Value
            Range("Q3").Value = Cells(i, 12).Value
        ElseIf Cells(i, 13).Value = Application.WorksheetFunction.Max(Range("M2:M" & LastRow)) Then
            Range("P4").Value = Cells(i, 10).Value
            Range("Q4").Value = Cells(i, 13).Value
        End If
    Next i
    
    'formatting
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "0"
    Columns("A:Q").AutoFit
    

End Sub
