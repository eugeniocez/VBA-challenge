# VBA-challenge

## This is the README file for my Module 2 Challenge

### During this challenge, I was able to complete the code by reviewing exercises done during previous classes. I also helped myself by making use of the Xpert Learning Assistant, ChatGPT and received 1 on 1 help during a tutoring session.

### Some of the help I received by making use of AI was in this section of my code, primarily because I when I ran the code, I was not getting the correct yearly change and percentage change.

`For i = 2 To RowCount
        '   Check if the next row has a different ticker symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Total = Total + ws.Cells(i, 7).Value
            '   Calculation for yearly change
            yearlyChange = ws.Cells(i, 6).Value - openingPrice
            '   Updating opening price for nect ticker
            openingPrice = ws.Cells(i + 1, 3).Value
            '   Calculation for percentage change
            percentageChange = (ws.Cells(i, 6) - ws.Cells(start, 3)) / ws.Cells(start, 3)`

### Then during my tutor session, my tutor helped me a lot with many sections of my code, he helped overall to make the code look cleaner and organized with the appropriate indents, and other sections of the code, as I couldn't figure out how to work out adding number formats and functionality to my script.

`If Total = 0 Then
        ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
        ws.Range("J" & 2 + j).Value = 0
        ws.Range("K" & 2 + j).Value = "%" & 0
        ws.Range("L" & 2 + j).Value = 0
            Else`

`        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = yearlyChange
                ws.Range("J" & 2 + j).NumberFormat = "0.00"
                ws.Range("K" & 2 + j).Value = percentageChange
                ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                ws.Range("L" & 2 + j).Value = Total`


`        Code to accumulate total volume for the same ticker symbol
            Total = Total + ws.Cells(i, 7).Value
        End If
    Next i`

   
    `   Output greatest percentage increase, decrease, and total volume
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
    
    '   Find the row number of the greatest percentage increase, decrease, and total volume
    Dim increase_number As Double
    Dim decrease_number As Double
    Dim volume_number As Double
  
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
    
    '   Output ticker symbols for the greatest percentage increase, decrease, and total volume
    ws.Range("P2") = ws.Cells(increase_number + 1, 9)
    ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("P4") = ws.Cells(volume_number + 1, 9)`

