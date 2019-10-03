Sub Stocks()

'Create sheet variables
Dim columncount, outputrow As Integer

'Create worksheet variable
Dim ws As Worksheet

'Create variables: op- opening price, oop - original opening price, cp - closing price
Dim op, oop, cp, vol, tvol, percent As Double
Dim columnend, rowend As String

tvol = 0


'Looping through all notebooks
For Each ws In Worksheets

    'Set Output row
    outputrow = 2

    'Find ranges for data in sheet
    rowend = ws.Cells(Rows.Count, 1).End(xlUp).Row
    columnend = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Print headers in each worksheet
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    
    'Looping through each row
    For i = 2 To rowend
        vol = 0
        vol = ws.Cells(i, 7)
        tvol = tvol + vol
    
        ' Storing opening and closing prices information
        cp = ws.Cells(i, 6)
        op = ws.Cells(i, 3)
        
        'Saving the original opening price
        If ws.Cells(i, 2).Value < ws.Cells(i - 1, 2).Value Then
            oop = op
            
        'Checking to see if we are looking at a new stock
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Output ticker
            ws.Cells(outputrow, 9) = ws.Cells(i, 1)
            
            'Calculate and store difference between year open and year close
            ws.Cells(outputrow, 10) = (oop - cp)
            If ws.Cells(outputrow, 10).Value >= 0 Then
                ws.Cells(outputrow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(outputrow, 10).Interior.ColorIndex = 4
            End If
            
            'Calculate Percentage Loss/Gain
            If oop <> 0 Then
                percent = ((oop - cp) / oop)
            Else
                percent = 0
            End If
        
            ws.Cells(outputrow, 11) = percent
            ws.Cells(outputrow, 11).NumberFormat = "0.00%"
            ' Print Total Volume
            ws.Cells(outputrow, 12) = tvol
            
            'Reset variables
            tvol = 0
            outputrow = outputrow + 1
            
        End If
    Next i
Next

End Sub

