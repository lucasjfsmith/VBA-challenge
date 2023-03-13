Sub summarizeStocks()
    
    'create a for loop to go through each worksheet
    For Each ws In Worksheets
    
        'create the summary columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'declare variables
        Dim i, lastRow, nextRow, prevRow, summaryRow As Integer
        Dim yrOpenPrice, yrClosePrice, volTotal As Double
    
        'find last row in dataset
        lastRow = ws.Cells(1, 1).End(xlDown).Row
    
        'initialize the volume counter to zero and the summary row so it begins in row 2
        volTotal = 0
        summaryRow = 2
    
        'create a for loop to go through each row in the worksheet
        For i = 2 To lastRow
        
            'set the surrounding row values
            nextRow = i + 1
            prevRow = i - 1
            
            'increment the volume total
            volTotal = volTotal + ws.Cells(i, 7).Value

            'check if current row matches previous row
            If ws.Cells(i, 1).Value <> ws.Cells(prevRow, 1) Then
                
                'record the opening price for the year
                yrOpenPrice = ws.Cells(i, 3).Value
                
            End If
            
            'check if the current row matches next row value
            If ws.Cells(i, 1).Value <> ws.Cells(nextRow, 1) Then
                
                'record the closing price for the year
                yrClosePrice = ws.Cells(i, 6).Value
                
                'populate summary values
                ws.Cells(summaryRow, 9).Value = Cells(i, 1).Value 'Ticker
                ws.Cells(summaryRow, 10).Value = yrClosePrice - yrOpenPrice 'Yearly Change
                ws.Cells(summaryRow, 11).Value = (yrClosePrice - yrOpenPrice) / yrOpenPrice 'Percent Change
                ws.Cells(summaryRow, 12).Value = volTotal
                
                'FORMAT CELLS
                'highlight Yearly Chang based on positive or negative value
                
                If ws.Cells(summaryRow, 10) > 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4 'green
                ElseIf ws.Cells(summaryRow, 10) < 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3 'red
                End If
                
                'format Percent Change as percentage rounded to two decimal places
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                ws.Cells(summaryRow, 10).NumberFormat = "$0.00"
                
                'reset volume total variable
                volTotal = 0
                
                'increase summary row for next summary
                summaryRow = summaryRow + 1
                
            End If
            
        Next i
        
        'auto fit sumamry table after complete
        ws.Range("I:L").Columns.AutoFit
        
        'CREATE LARGEST CHANGE TABLE
        'format rows and columns
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'create variables for tracking largest changes ticker symbols
        Dim maxPercentIncreaseTicker, maxPercentDecreaseTicker, maxTotalVolumeTicker As String
        
        'create variables for current max values
        Dim maxPercentIncrease, maxPercentDecrease, maxTotalVolume As Double
        
        'initialize values
        maxPercentIncrease = 0
        maxPercentDecrease = 0
        maxTotalVolume = 0
        
        'find last row in summary table
        Dim lastSumRow As Integer
        lastSumRow = ws.Cells(1, 9).End(xlDown).Row
        
        'create for loop to find max values
        For i = 2 To lastSumRow
        
            'check if percent change greater than current max increase
            If ws.Cells(i, 11).Value > maxPercentIncrease Then
                
                'update the max percent variable and row tracker
                maxPercentIncrease = ws.Cells(i, 11).Value
                maxPercentIncreaseTicker = ws.Cells(i, 9).Value
            
            End If
            
            'check if percent change less than current max decrease
            If ws.Cells(i, 11).Value < maxPercentDecrease Then
                
                'update the max percent variable and row tracker
                maxPercentDecrease = ws.Cells(i, 11).Value
                maxPercentDecreaseTicker = ws.Cells(i, 9).Value
            
            End If
            
            'check if total volume greater than current max total volume
            If ws.Cells(i, 12).Value > maxTotalVolume Then
                
                'update the max percent variable and row tracker
                maxTotalVolume = ws.Cells(i, 12).Value
                maxTotalVolumeTicker = ws.Cells(i, 9).Value
            
            End If
            
        Next i
            
        'populate max percent increase
        ws.Range("P2").Value = maxPercentIncreaseTicker
        ws.Range("Q2").Value = maxPercentIncrease
        
        'populate max percent decrease
        ws.Range("P3").Value = maxPercentDecreaseTicker
        ws.Range("Q3").Value = maxPercentDecrease
        
        'populate max total volume
        ws.Range("P4").Value = maxTotalVolumeTicker
        ws.Range("Q4").Value = maxTotalVolume
        
        'format percentages
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'autofit the table
        ws.Range("O1:Q4").Columns.AutoFit
        
                
    Next ws
    
End Sub

Sub clear() 'used to clear data from spreadsheet when testing
    
    For Each ws In Worksheets
    
        ws.Range("I:Q").ClearContents
        ws.Range("I:Q").ClearFormats
    
    Next ws

End Sub

