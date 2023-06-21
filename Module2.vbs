Attribute VB_Name = "Module2"
Sub Run()
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    'Define variables we will use
    Dim table_row As Double
    Dim ticker As String
    Dim volume As Double
    Dim lastrow As Double
    Dim year_start As Double
    Dim year_end As Double
    Dim year_change As Double
    Dim year_percent As Double
    
    
    'Define inital values of  some variables
    table_row = 2
    volume = 0
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Define title cells
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    year_start = ws.Cells(2, 3)
    'Time to loop through the stock
    For i = 2 To lastrow
    
        'Check if next cell is different as an initial parameter
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'If it is not, we need to assign the appropriate values to the Cells
            ticker = ws.Cells(i, 1).Value
            ws.Range("I" & table_row).Value = ticker
            ws.Range("L" & table_row).Value = volume
            
            'We also need to calculate the year change and store it in the cells
            year_end = ws.Cells(i, 6)
            year_change = year_end - year_start
            ws.Range("J" & table_row).Value = year_change
            
                'Then color code it
                If ws.Range("J" & table_row).Value > 0 Then
                    ws.Range("J" & table_row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & table_row).Interior.ColorIndex = 3
                End If
            
            'Have to be careful with the percentage as we can't divide by 0
            year_percent = year_end / year_start - 1
                If IsError(year_percent) Then
                    year_percent = 0
                End If
            
            'Assign rest of values and change type
            ws.Range("K" & table_row).Value = year_percent
            ws.Range("K" & table_row).NumberFormat = "0.00%"
            
            'Now we reset our values to prepare for next loop
            year_start = ws.Cells(i + 1, 3).Value
            volume = 0
            table_row = table_row + 1
        Else
            'If they are the same, then we just want to add the volume
            volume = volume + ws.Cells(i, 7)
        End If
    Next i
    
    'Let's do the leaders now
    'Create the text cells
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'And the variables, I'll set volume up as a different cell so no need to repeat
    Dim highest As Double
    Dim lowest As Double
    
    'Assign them starting cell values
    highest = ws.Cells(2, 11).Value
    lowest = ws.Cells(2, 11).Value
    volume = ws.Cells(2, 12).Value
    
    high_ticker = ws.Cells(2, 9).Value
    low_ticker = ws.Cells(2, 9).Value
    volume_ticker = ws.Cells(2, 9).Value
    
    
    'Time to run a for loop
    'Because I have stored values above, I can just check the next cell
    'If the next cell is better suited for what I'm looking for
    'then I replace the stored value
    
    For i = 2 To lastrow
        If ws.Cells(i, 11) > highest Then
            highest = ws.Cells(i, 11).Value
            high_ticker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 11) < lowest Then
            lowest = ws.Cells(i, 11).Value
            low_ticker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 12) > volume Then
            volume = ws.Cells(i, 12).Value
            volume_ticker = ws.Cells(i, 9).Value
        End If
    Next i
    
    'Then print out the rest
    ws.Cells(2, 17).Value = highest
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(2, 16).Value = high_ticker
    ws.Cells(3, 17).Value = lowest
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = low_ticker
    ws.Cells(4, 17).Value = volume
    ws.Cells(4, 16).Value = volume_ticker
Next ws
End Sub

