' UofT SCS EdX Data Bootcamp: Challenge 2
' Script Author: Tania Barrera (tsbarr)
' ---
' ** Run Main() for final solution **
' ---
' Assumptions:
        ' each sheet contains data from a single year
        ' data from the year is organized as daily data, with a row containing data for a single stock ticker and a single date
        ' data always starts at the top left corner of the sheet, has 1 header row and contains 7 columns:
                ' 1     <ticker>      ticker symbol
                ' 2     <date>        date of the daily data in that row
                ' 3     <open>        opening price for the date
                ' 4     <high>        high price for the date, not used
                ' 5     <low>         low price for the date, not used
                ' 6     <close>       closing price for the date
                ' 7     <vol>         volume of stock for the date
        ' there are no blanks in the <ticker> column, the first blank from the bottom of the column will be taken as the end of the data

' -----------------------------------------------------------
' MAIN SUB PROCEDURE (ALL WORKSHEETS ANALYSIS)
' -----------------------------------------------------------
' --- Main sub loops through sheets and run the CalculateYearData() Sub
' Based on source: https://excelchamps.com/vba/loop-sheets/
Sub Main():
	Dim ws As Worksheet
	For Each ws In ThisWorkbook.Worksheets
	ws.Activate ' activate ws first, so below sub applies to it
	CalculateYearData
	Next ws
End Sub
' -----------------------------------------------------------
' SINGLE WORKSHEET ANALYSIS - called by main to loop through all sheets
' -----------------------------------------------------------
' --- Sub CalculateYearData runs all required analyses in a **single** sheet
Sub CalculateYearData():
        ' Turn off screen updating to improve performance, as per:
        ' https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/fill-a-value-down-into-blank-cells-in-a-column
        Application.ScreenUpdating = False

        ' --- variable declaration
	' % is short for "as integer"
	' $ is short for "as string"
	' & is short for for "as long"
	' # is short for "as double"
        Dim lastRow& ' to id last row with data
        Dim inputArray() As Variant, resultArray() As Variant ' type variant so we can use different data types
        Dim uniqueTickers As New Collection ' to find all unique ticker names
        ' for use within loops
        Dim thisTicker$, thisDate&, thisOpen#, thisClose#, thisVol As LongLong
	Dim i&, j& ' counters
        ' resultArray column index integer values, to improve readability
        Dim tickerI%, firstDateI%, firstOpenI%, lastDateI%, lastCloseI%, yearChangeI%, percentChangeI%, totalVolI%
        ' max and min variables
        Dim minPercentChangeName$, minPercentChangeValue#
        Dim maxPercentChangeName$, maxPercentChangeValue#
        Dim maxTotalVolName$, maxTotalVolValue As LongLong
        
        ' --- read data into inputArray
        ' Find last row with data in <ticker> column (1)
        ' Source: https://www.excelcampus.com/vba/find-last-row-column-cell/
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        ' use lastRow to read in range of input data, no headers
        inputArray = Range("A2:G" & lastRow)
	' this improves performance, since there is no need to keep reading data from worksheet, as per:
	' https://www.soa.org/news-and-publications/newsletters/compact/2012/january/com-2012-iss42/excel-vba-speed-and-efficiency/

        ' --- get unique ticker names
        ' based on: https://stackoverflow.com/a/3017973
	' use error-handling: if ticker name exists as a key in the collection, it will throw error, so it will skip it
        On Error Resume Next 
	' loop through all rows in inputArray
	' note: inputArray is 1-based because it comes from a Range object
        For i = 1 To UBound(inputArray, 1)
		thisTicker = inputArray(i, 1) ' get ticker in current inputArray row
		uniqueTickers.Add thisTicker, thisTicker ' add to uniqueTickers collection using a key, to avoid repeats
        Next i
        On Error GoTo 0 ' disable error-handling

        ' --- initialize resultArray
        ' use number of uniqueTickers to declare resultArray, since we now know the number of rows we need
        ' we need an array and not collection so we can perform operations, as per: https://excelmacromastery.com/excel-vba-collections/
        ' see here for declaring dynamic arrays:
        ' https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/declaring-arrays

        ' set array dimensions
	' note: array dimensions are 0-based, but collection counts are 1-based, so we use .Count-1
        ReDim resultArray(uniqueTickers.Count - 1, 7) 
        ' resultArray has 8 columns
        ' use column index variables to make writing/reading code easier
        tickerI = 0
        firstDateI = 1
        firstOpenI = 2
        lastDateI = 3
        lastCloseI = 4
        yearChangeI = 5
        percentChangeI = 6
        totalVolI = 7

        ' --- write uniqueTickers values to first column of resultArray (tickerI = 0)
        For i = 1 To uniqueTickers.Count ' loop through uniqueTickers collection
        	resultArray(i - 1, tickerI) = uniqueTickers(i) ' note: resultArray is 0-based, but collections are 1-based
        Next i

        ' --- find firstDate, firstOpen, lastDate, lastClose and totalVol
        ' iterate within inputArray, fill values in resultArray
        ' note: inputArray is 1-based because it comes from a Range object
        For i = 1 To UBound(inputArray, 1)
		' get data of current row
                thisTicker = inputArray(i, 1)
                thisDate = inputArray(i, 2)
                thisOpen = inputArray(i, 3)
                thisClose = inputArray(i, 6)
                thisVol = inputArray(i, 7)
                ' find row index of thisTicker in resultArray
                ' source for search within array: https://www.excel-pratique.com/en/vba_tricks/search-in-array-function
                For j = 0 To UBound(resultArray, 1)
                        If resultArray(j, tickerI) = thisTicker Then 'If value found
                                ' current row j is index of thisTicker in resultArray 
                                Exit For
                        End If
                Next j
                ' check if this is the first time we find thisTicker by checking if value firstDate is empty
                If IsEmpty(resultArray(j, firstDateI)) Then
                        ' initialize with current row info
                        resultArray(j, firstDateI) = thisDate    ' firstDate
                        resultArray(j, firstOpenI) = thisOpen    ' firstOpen
                        resultArray(j, lastDateI) = thisDate     ' lastDate
                        resultArray(j, lastCloseI) = thisClose   ' lastClose
                        resultArray(j, totalVolI) = thisVol      ' totalVol
                Else ' we have seen thisTicker before
                        ' check if this date is earlier than firstDate
                        If thisDate < resultArray(j, firstDateI) Then
                        ' replace firstDate with thisDate
                        resultArray(j, firstDateI) = thisDate
                        ' and firstOpen with thisOpen
                        resultArray(j, firstOpenI) = thisOpen
                        End If
                        ' check if thisDate is later than lastDate
                        If thisDate > resultArray(j, lastDateI) Then
                        ' replace lastDate with thisDate
                        resultArray(j, lastDateI) = thisDate
                        ' and lastClose with thisClose
                        resultArray(j, lastCloseI) = thisClose
                        End If
                        ' add thisVol to totalVol
                        resultArray(j, totalVolI) = resultArray(j, totalVolI) + thisVol
                End If
        Next i

	' --- year data labels
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"

        ' --- initialize max and min variables
        minPercentChangeName = ""
        minPercentChangeValue = 0
        maxPercentChangeName = ""
        maxPercentChangeValue = 0
        maxTotalVolName = ""
        maxTotalVolValue = 0

        ' --- calculate year and percent change, write out year data and get min/max variables
        ' iterate within resultArray
        For i = 0 To UBound(resultArray, 1)
                ' * calculate year and percent change *
                ' yearChange = lastClose - firstOpen
                resultArray(i, yearChangeI) = resultArray(i, lastCloseI) - resultArray(i, firstOpenI)
                ' percentChange = yearChange / firstOpen
                resultArray(i, percentChangeI) = resultArray(i, yearChangeI) / resultArray(i, firstOpenI)

                ' ** write columns of interest to spreadsheet **
                ' column ticker to column I (9)
                Cells(i + 2, 9) = resultArray(i, tickerI)
                ' column yearChange to column J (10)
                Cells(i + 2, 10) = resultArray(i, yearChangeI)
                ' * conditional formatting of yearChange *
                With Cells(i + 2, 10).Interior ' format interior color of cell
                        If resultArray(i, yearChangeI) > 0 Then ' positive
                                .Color = RGB(78, 163, 54) ' green
                        ElseIf resultArray(i, yearChangeI) < 0 Then ' negative
                                .Color = RGB(204, 39, 39) ' red
                        Else ' it must be equal to 0
                                .ColorIndex = 44 ' yellow
                        End If
                 End With
                ' column percentChange to column K (11), with percent format
                Cells(i + 2, 11) = FormatPercent(resultArray(i, percentChangeI))
                ' * conditional formatting of percentChange *
                With Cells(i + 2, 11).Interior ' format interior color of cell
                        If resultArray(i, percentChangeI) > 0 Then ' positive
                                .Color = RGB(78, 163, 54) ' green
                        ElseIf resultArray(i, percentChangeI) < 0 Then ' negative
                                .Color = RGB(204, 39, 39) ' red
                        Else ' it must be equal to 0
                                .ColorIndex = 44 ' yellow
                        End If
                 End With
                ' column totalVol to column L (12)
                Cells(i + 2, 12) = resultArray(i, totalVolI)

		' * find min/max variables *
                ' compare to get minPercentChange
                If resultArray(i, percentChangeI) < minPercentChangeValue Then
                        minPercentChangeValue = resultArray(i, percentChangeI)
                        minPercentChangeName = resultArray(i, tickerI)
                End If
                ' compare to get maxPercentChange
                If resultArray(i, percentChangeI) > maxPercentChangeValue Then
                        maxPercentChangeValue = resultArray(i, percentChangeI)
                        maxPercentChangeName = resultArray(i, tickerI)
                End If
                ' compare to get maxTotalVol
                If resultArray(i, totalVolI) > maxTotalVolValue Then
                        maxTotalVolValue = resultArray(i, totalVolI)
                        maxTotalVolName = resultArray(i, tickerI)
                End If
        Next i

        ' --- output min/max variables
        ' row labels to col 15
        Cells(2, 15) = "Greatest % Increase"
        Cells(3, 15) = "Greatest % Decrease"
        Cells(4, 15) = "Greatest Total Volume"
        
        ' ticker names to col 16
        Cells(1, 16) = "Ticker" ' column label
        Cells(2, 16) = maxPercentChangeName 	' greatest % increase, ticker name
        Cells(3, 16) = minPercentChangeName 	' greatest % decrease, ticker name
        Cells(4, 16) = maxTotalVolName 		' greatest total volume, ticker name
        ' values to col 17, max and min percent changes with percent format
        Cells(1, 17) = "Value" ' column label
        Cells(2, 17) = FormatPercent(maxPercentChangeValue)	' greatest % increase, percent format
        Cells(3, 17) = FormatPercent(minPercentChangeValue)	' greatest % decrease, percent format
        Cells(4, 17) = maxTotalVolValue				' greatest total volume

        ' --- turn screen updating back on
        Application.ScreenUpdating = True

End Sub
