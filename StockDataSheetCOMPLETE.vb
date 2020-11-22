Sub StockDataSheet()
'Setting each sheet as a indivudal within the whole Excel document
Dim ws As Worksheet
    'Start loop through each individual sheet
    For Each ws In Worksheets

        'Setting column labels for assignment
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'Setting column and row labels for bonus
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        'Dim variables needed throughout the whole script
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim StockVolume As Variant: StockVolume = 0
        Dim MaxPercent As Double: MaxPercent = 0
        Dim MinPercent As Double: MinPercent = 0
        Dim MaxVolume As Variant: MaxVolume = 0
        Dim MaxTicker As String
        Dim MinTicker As String
        Dim MaxVolTicker As String
        
        'Dim variables use to define the rows for each piece of info
        Dim TickerRow As Long: TickerRow = 1
        Dim ChangeRow As Long: ChangeRow = 1
        Dim PercentRow As Long: PercentRow = 1
        Dim VolumeRow As Long: VolumeRow = 1

        'Setting variable to use and find last row of data on the sheet. Needs to be a Long because data set is so large
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        'Setting up for the needed For loop. i needs to be long to match the LastRow value since the data is so large
        Dim i As Long
    
        'Set starting OpeningPrice
        OpeningPrice = ws.Cells(2, 3).Value
    
            'Start For loop to loop through all the rows. Nested If statement within loop
            For i = 2 To LastRow
                'If statement that looks at if the ticker matches the ticker in the cell below.
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    'Adding to the rows so it moves down to print the next ticker's info
                    TickerRow = TickerRow + 1
                    ChangeRow = ChangeRow + 1
                    PercentRow = PercentRow + 1
                    VolumeRow = VolumeRow + 1
                    'Finds end of individaul tickers and prints that ticker symbol
                    Ticker = ws.Cells(i, 1).Value
                    ws.Cells(TickerRow, 9).Value = Ticker
                    'Calculates both yearly change and percent change. Prints info into cell. Nest If statments to look at specifc criteria
                    ClosingPrice = ws.Cells(i, 6).Value
                    YearlyChange = ClosingPrice - OpeningPrice
                    ws.Cells(ChangeRow, 10).Value = YearlyChange
                        If OpeningPrice = 0 Then
                            ws.Cells(PercentRow, 11).Value = 0
                            OpeningPrice = ws.Cells(i + 1, 3).Value
                        Else
                            PercentChange = (YearlyChange / OpeningPrice) * 100
                            PercentChange = Round(PercentChange, 2)
                            ws.Cells(PercentRow, 11).Value = (CStr(PercentChange) & "%")
                            OpeningPrice = ws.Cells(i + 1, 3).Value
                        End If
                    'Totals stock volume based on the If statment finding the end of the indivudal ticker's info
                    StockVolume = StockVolume + ws.Cells(i, 7).Value
                    ws.Cells(VolumeRow, 12).Value = StockVolume
                        'Nested If statements to find the information for the bonus
                        If (PercentChange > MaxPercent) Then
                                MaxPercent = PercentChange
                                MaxTicker = Ticker
                        ElseIf (PercentChange < MinPercent) Then
                                MinPercent = PercentChange
                                MinTicker = Ticker
                        End If
                           
                        If (StockVolume > MaxVolume) Then
                            MaxVolume = StockVolume
                            MaxVolTicker = Ticker
                        End If
                    
                    'Resetting the variables for the next loop through
                    PercentChange = 0
                    StockVolume = 0
                'Other half of If statement that looks to see if the ticker symbols match. If they do, the stock volume will increase
                ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                        StockVolume = StockVolume + ws.Cells(i, 7).Value
                End If
                'If statement to run through and color the cells based on the change is negative or positive
                If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                End If
            'Go to the next i in the For loop
            Next i
        'Printing values for the bonus to the correct cells
        ws.Cells(2, 17).Value = (CStr(MaxPercent) & "%")
        ws.Cells(3, 17).Value = (CStr(MinPercent) & "%")
        ws.Cells(2, 16).Value = MaxTicker
        ws.Cells(3, 16).Value = MinTicker
        ws.Cells(4, 17).Value = MaxVolume
        ws.Cells(4, 16).Value = MaxVolTicker
    'Go to the next sheet and run through whole script again
    Next ws
End Sub
