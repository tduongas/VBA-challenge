Sub Stock_Analyzer_Part_1()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' Created a Variable to Hold WorksheetName
        Dim WorksheetName As String

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Clear report area
        ws.Range("$I$1:$T$" & LastRow).Clear
        
        ' Determine the Last Column Number
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).column
        
        reportCellRow = 1
        ReportCellColumn = 9
        
        ' create reportHeadings
        ws.Cells(reportCellRow, ReportCellColumn).Value = "<ticker>"
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = "<opening_price>"
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = "<closing_price>"
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = "<yearly_change>"
        ws.Cells(reportCellRow, ReportCellColumn).EntireColumn.EntireColumn.AutoFit
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = "<yearly_percentage_change>"
        ws.Cells(reportCellRow, ReportCellColumn).EntireColumn.EntireColumn.AutoFit
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = "<yearly_stock_volume>"
        ws.Cells(reportCellRow, ReportCellColumn).EntireColumn.EntireColumn.AutoFit
        
        ' set column row column of reporting fields to cell I2
        ReportCellColumn = 9
        reportCellRow = 2
        
        ' Grab the WorksheetName
        ' MsgBox ("Worksheet: " & WorksheetName & ", LastRow: " & LastRow & ", LastColumn: " & LastColumn)
        WorksheetName = ws.Name

        ' Set a variable for specifying the column of interest
        Dim column As Integer
        column = 1
            
        Dim tickerItem As String
        Dim newTickerRow As Long
        
        Dim yearlyOpeningPrice As Double
        Dim yearlyClosingPrice As Double
        Dim yearlyChange As Double
        Dim yearlyPercentageChange As Double
        Dim yearlyStockVolume As Variant
        
        
        ' start at row 2, after the heading information
        newTickerRow = 2
            
        ' Loop through rows in the column
        For i = 2 To LastRow
        
            ' Loop through the columns to get opening price
            For J = 1 To LastColumn
    
                If i = newTickerRow And J = 3 Then
                
                    ' MsgBox ("Loop 1 Row: " & i & " Column: " & J & " | " & ws.Cells(i, J).Value)
                    yearlyOpeningPrice = ws.Cells(i, J).Value
                
                ElseIf J = LastColumn Then
                    
                    ' MsgBox ("Loop 1 Row: " & i & " Column: " & J & " | " & ws.Cells(i, J).Value)
                    yearlyStockVolume = CDec(yearlyStockVolume) + CDec(ws.Cells(i, J).Value)
                    
                End If
                
            Next J
            
            ' Searches for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
        
                ' Message Box the value of the current cell and value of the next cell
                ' MsgBox (ws.Cells(i, column).Value & " and then " & ws.Cells(i + 1, column).Value)
                
                ' Store the ticketItem, MsgBox ("Ticker: " & ws.Cells(i, column).Value)
                tickerItem = ws.Cells(i, column).Value
                
                ' Loop through the columns at the closing price
                For k = 1 To LastColumn
    
                    If k = 6 Then
                    
                        ' MsgBox ("Loop 2 Row: " & i & " Column: " & J & " | " & ws.Cells(i, J).Value)
                        yearlyClosingPrice = ws.Cells(i, k).Value
                        
                        ' Calculate yearly change
                        yearlyChange = yearlyClosingPrice - yearlyOpeningPrice
                         
                        ' Cater for division by 0
                        If yearlyOpeningPrice = 0 And yearlyChange = 0 Then
                            yearlyPercentageChange = 0
                        Else
                            ' Calculate yearly percentage change
                            yearlyPercentageChange = (Round(yearlyChange, 3) / Round(yearlyOpeningPrice, 3))
                        End If
                                                                        
                        ' MsgBox ("Ticker: " & tickerItem & ", yearlyOpen: " & yearlyOpeningPrice & ", yearlyClose: " & yearlyClosingPrice & ", yearlyChange: " & yearlyChange & ", yearlyPercentageChange: " & yearlyPercentageChange & ", yearlyStockVolume: " & yearlyStockVolume)
                        
                        ' Inserting Data Via Cells
                        ws.Cells(reportCellRow, ReportCellColumn).Value = tickerItem
                        ReportCellColumn = ReportCellColumn + 1
                        
                        ws.Cells(reportCellRow, ReportCellColumn).Value = yearlyOpeningPrice
                        ReportCellColumn = ReportCellColumn + 1
                        
                        ws.Cells(reportCellRow, ReportCellColumn).Value = yearlyClosingPrice
                        ReportCellColumn = ReportCellColumn + 1
                        
                        ws.Cells(reportCellRow, ReportCellColumn).Value = Round(yearlyChange, 2)
                        
                        If yearlyChange > 0 Then
                            ws.Cells(reportCellRow, ReportCellColumn).Interior.ColorIndex = 4
                        Else
                            ws.Cells(reportCellRow, ReportCellColumn).Interior.ColorIndex = 3
                        End If
                        
                        ReportCellColumn = ReportCellColumn + 1
                        ws.Cells(reportCellRow, ReportCellColumn).Value = Round(yearlyPercentageChange, 3)
                        ws.Cells(reportCellRow, ReportCellColumn).NumberFormat = "0.00%"
                        ReportCellColumn = ReportCellColumn + 1
                        ws.Cells(reportCellRow, ReportCellColumn).Value = yearlyStockVolume
                        
                        ' go down one line on new ticker report
                        reportCellRow = reportCellRow + 1
                        
                        ' reset the report cell column back to cell I or cell 9
                        ReportCellColumn = 9

                        ' Reset the yearlyStockVolume
                        yearlyStockVolume = 0
                        
                    End If
                
                Next k
                
                ' MsgBox ("Row: " & i)
                newTickerRow = i + 1
            
            End If
            
        Next i
        
        ' If WorksheetName = "A" Then Exit For
        
    Next ws

End Sub


Sub Stock_Analyzer_Part_2()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' Created a Variable to Hold WorksheetName
        Dim WorksheetName As String
        Dim maximumPercentageValue As Double

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Clear report area
        ws.Range("$O$1:$T$" & LastRow).Clear
        
       
        ' Determine last row in selected column
        reportLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        MsgBox ("LastRow: " & reportLastRow)
        
        reportCellRow = 1
        ReportCellColumn = 16
        
        ' Create reportHeadings
        ws.Cells(reportCellRow, ReportCellColumn).Value = ""
        ws.Cells(reportCellRow, ReportCellColumn).EntireColumn.EntireColumn.AutoFit
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = "<ticker>"
        ws.Cells(reportCellRow, ReportCellColumn).EntireColumn.EntireColumn.AutoFit
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = "<value>"
        ws.Cells(reportCellRow, ReportCellColumn).EntireColumn.EntireColumn.AutoFit
        ReportCellColumn = ReportCellColumn + 1
        
        ' Set column row column of reporting fields to cell P2
        reportCellRow = 2
        ReportCellColumn = 16
        
        ' Grab the WorksheetName
        WorksheetName = ws.Name
        
        
        ' ------------------------------
        ' GREATEST PERCENTAGE INCREASE
        ' ------------------------------
        
        ' Get maximum percentage value
        maximumPercentageValue = WorksheetFunction.Max(Range("M2:M" & reportLastRow))
        
        ' Get the maximum percentage row
        maximumPercentageRow = WorksheetFunction.Match(maximumPercentageValue, Range(("M2:M" & reportLastRow)), 0) + Range(("M2:M" & reportLastRow)).Row - 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = "Greatest % Increase"
        ws.Cells(reportCellRow, ReportCellColumn).EntireColumn.EntireColumn.AutoFit
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = Range("I" & maximumPercentageRow & ":I" & maximumPercentageRow).Value
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = maximumPercentageValue
        ws.Cells(reportCellRow, ReportCellColumn).NumberFormat = "0.00%"
        
        ' Offset to next line
        reportCellRow = 3
        ReportCellColumn = 16
        
        
        ' ------------------------------
        ' GREATEST PERCENTAGE DECREASE
        ' ------------------------------
        
        ' Get the min percentage value
        minimumPercentageValue = WorksheetFunction.Min(ws.Range("M2:M" & reportLastRow))
        
        ' Get the maximum percentage row
        minimumPercentageRow = WorksheetFunction.Match(minimumPercentageValue, ws.Range(("M2:M" & reportLastRow)), 0) + ws.Range(("M2:M" & reportLastRow)).Row - 1
        
        ' Output results
        ws.Cells(reportCellRow, ReportCellColumn).Value = "Greatest % Decrease"
        ws.Cells(reportCellRow, ReportCellColumn).EntireColumn.EntireColumn.AutoFit
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = ws.Range("I" & minimumPercentageRow & ":I" & minimumPercentageRow).Value
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = minimumPercentageValue
        ws.Cells(reportCellRow, ReportCellColumn).NumberFormat = "0.00%"
        
        ' Offset to next line
        reportCellRow = 4
        ReportCellColumn = 16
        
        
        ' ------------------------------
        ' GREATEST TOTAL VOLUME
        ' ------------------------------
        
        ' Get the min percentage value
        maximumTotalValue = WorksheetFunction.Max(ws.Range("N2:N" & reportLastRow))
        
        ' Get the maximum percentage row
        maximumTotalVolumeRow = WorksheetFunction.Match(maximumTotalValue, ws.Range(("N2:N" & reportLastRow)), 0) + ws.Range(("N2:N" & reportLastRow)).Row - 1
               
        ' Output results
        ws.Cells(reportCellRow, ReportCellColumn).Value = "Greatest % Decrease"
        ws.Cells(reportCellRow, ReportCellColumn).EntireColumn.EntireColumn.AutoFit
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = ws.Range("I" & maximumTotalVolumeRow & ":I" & maximumTotalVolumeRow).Value
        ReportCellColumn = ReportCellColumn + 1
        
        ws.Cells(reportCellRow, ReportCellColumn).Value = maximumTotalValue

        ' Exit if worksheetName  = "A"
        ' If WorksheetName = "A" Then Exit For
        
    Next ws


End Sub