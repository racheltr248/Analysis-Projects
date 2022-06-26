Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
        
    Worksheets("2018").Activate
    
    'Set initial volume to zero
    totalVolume = 0
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'Establish the number of rows to loop over
    rowStart = 2
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    
    'Loop over all the rows
    For i = rowStart To rowEnd
    
        'increase totalVolume if ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
            'Increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value
        End If
        
        'startingPrice
        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
            startingPrice = Cells(i, 6).Value
        End If
        
        'endingPrice
        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
            endingPrice = Cells(i, 6).Value
        End If
        
    Next i
       
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1
    
End Sub


Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer

'1. Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

'2. Initialize an array of all tickers.
    Dim tickers(11) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

'3. Prepare for the analysis of tickers.

    '3a. Initialize variables for the starting price and ending price.
        Dim startingPrice As Double
        Dim endingPrice As Double

    '3b. Activate the data worksheet.
        Worksheets(yearValue).Activate

    '3c. Find the number of rows to loop over.
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
'4. Loop through the tickers.
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0

'5. Loop through rows in the data.
        Worksheets(yearValue).Activate
        For j = 2 To RowCount

    '5a. Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If

    '5b. Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
    '5c. Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        Next j

'6. Output the data for the current ticker.
    Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

Sub formatAllStocksAnalysisTable()
    'Activate the proper sheet for formatting
    Worksheets("All Stocks Analysis").Activate
    
    'Static formatting
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("A:C").AutoFit
    
    'Conditional color formatting with loop
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
        Else
            Cells(i, 3).Interior.Color = xlNone
        End If
    Next i
    
End Sub

Sub ClearWorksheet()
    Cells.Clear
End Sub

