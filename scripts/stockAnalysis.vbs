' UofT SCS EdX Data Bootcamp: Challenge 2
' Script Author: Tania Barrera (tsbarr)
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
' ---

Sub CalculateYearData():
	' Turn off screen updating to improve performance, as per:
	' https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/fill-a-value-down-into-blank-cells-in-a-column
	Application.ScreenUpdating = False

	' --- variable declaration
        Dim lastRow As Long
        Dim inputArray() As Variant, resultArray() As Variant ' type variant so we can use different data types
        Dim uniqueTickers As New Collection
        ' for within loops, $ is short for "as string", & is short for for "as long", # is short form for "as double"
        Dim thisTicker$, thisDate&, thisOpen#, thisClose#, thisVol As LongLong
        ' resultArray column index, % is short form for "as integer"
        Dim ticker%, firstDate%, firstOpen%, lastDate%, lastClose%, yearChange%, percentChange%, totalVol% ' integer for index
        ' counters / index within array iterators
        Dim i&, j&
	' max and min variables
	Dim minPercentChangeName$, minPercentChangeValue#
	Dim maxPercentChangeName$, maxPercentChangeValue#
	Dim maxTotalVolName$, maxTotalVolValue as LongLong
	
	' --- read data into inputArray
        ' Find last row with data in <ticker> column (1)
        ' Source: https://www.excelcampus.com/vba/find-last-row-column-cell/
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row

        ' use lastRow to read in range of input data, no headers
        inputArray = Range("A2:G" & lastRow)

        ' --- get unique ticker values
        ' based on source: https://stackoverflow.com/a/3017973
        On Error Resume Next ' if ticker is already in collection, skip, since it will throw error
        ' note: inputArray is 1-based because it comes from a Range object
        For i = 1 To UBound(inputArray, 1)
        thisTicker = inputArray(i, 1)
        uniqueTickers.Add thisTicker, thisTicker
        Next
        On Error GoTo 0

        ' --- initialize resultArray
	' use number of uniqueTickers to declare resultArray, since we now know the number of rows we need
        ' we need an array and not collection so we can perform operations, as per: https://excelmacromastery.com/excel-vba-collections/
        ' see here for declaring dynamic arrays:
        ' https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/declaring-arrays
        
        ReDim resultArray(uniqueTickers.Count - 1, 7) ' array dimensions, note: arrays are 0-based, but collections are 1-based
        ' resultArray has 8 columns
        ' column index in variables to make coding/reading easier
        ticker = 0
        firstDate = 1
        firstOpen = 2
        lastDate = 3
        lastClose = 4
        yearChange = 5
        percentChange = 6
        totalVol = 7

        ' --- write uniqueTickers to resultArray first column (ticker)
        For i = 1 To uniqueTickers.Count
        resultArray(i - 1, ticker) = uniqueTickers(i) ' note: resultArray is 0-based, but collections are 1-based
        Next i

	' --- find firstDate, firstOpen, lastDate, lastClose and totalVol
        ' iterate within inputArray, fill values in resultArray
        ' note: inputArray is 1-based because it comes from a Range object
        For i = 1 To UBound(inputArray, 1)
                thisTicker = inputArray(i, 1)
                thisDate = inputArray(i, 2)
                thisOpen = inputArray(i, 3)
                thisClose = inputArray(i, 6)
                thisVol = inputArray(i, 7)
                ' find row index of thisTicker in resultArray
                ' source for search within array: https://www.excel-pratique.com/en/vba_tricks/search-in-array-function
                For j = 0 To UBound(resultArray, 1)
                        If resultArray(j, ticker) = thisTicker Then 'If value found
                                ' current row j is index in resultArray of thisTicker
                                Exit For
                        End If
                Next j
                ' check if this is the first time we find thisTicker by checking if value firstDate is empty
                If IsEmpty(resultArray(j, firstDate)) Then
                        ' initialize with current row info
                        resultArray(j, firstDate) = thisDate    ' firstDate
                        resultArray(j, firstOpen) = thisOpen    ' firstOpen
                        resultArray(j, lastDate) = thisDate     ' lastDate
                        resultArray(j, lastClose) = thisClose   ' lastClose
                        resultArray(j, totalVol) = thisVol      ' totalVol
                Else ' we have seen thisTicker before
                        ' check if this date is earlier than firstDate
                        If thisDate < resultArray(j, firstDate) Then
                        ' replace firstDate with thisDate
                        resultArray(j, firstDate) = thisDate
                        ' and firstOpen with thisOpen
                        resultArray(j, firstOpen) = thisOpen
                        End If
                        ' check if thisDate is later than lastDate
                        If thisDate > resultArray(j, 3) Then
                        ' replace lastDate with thisDate
                        resultArray(j, lastDate) = thisDate
                        ' and lastClose with thisClose
                        resultArray(j, lastClose) = thisClose
                        End If
                        ' add thisVol to totalVol
                        resultArray(j, totalVol) = resultArray(j, totalVol) + thisVol
                End If
        Next i

	' --- initialize max and min variables
	minPercentChangeName = ""
	minPercentChangeValue = 0
	maxPercentChangeName = ""
	maxPercentChangeValue = 0
	maxTotalVolName = ""
	maxTotalVolValue = 0
        ' --- calculate year and percent change, write out and get min/max variables
	' iterate within resultArray
        For i = 0 To UBound(resultArray, 1)
		' calculate year and percent change
		' yearChange = lastClose - firstOpen
		resultArray(i, yearChange) = resultArray(i, lastClose) - resultArray(i, firstOpen)
		' percentChange = yearChange / firstOpen
		resultArray(i, percentChange) = resultArray(i, yearChange) / resultArray(i, firstOpen)
		' write to spreadsheet the columns of interest
		Cells(i + 2, 9) = resultArray(i, ticker)                ' column ticker to column I (9)
		Cells(i + 2, 10) = resultArray(i, yearChange)           ' column yearChange to column J (10)
		Cells(i + 2, 11) = resultArray(i, percentChange)        ' column percentChange to column K (11)
		Cells(i + 2, 12) = resultArray(i, totalVol)             ' column totalVol to column L (12)
		' compare to get minPercentChange
		If resultArray(i, percentChange) < minPercentChangeValue Then
			minPercentChangeValue = resultArray(i, percentChange)
			minPercentChangeName = resultArray(i, ticker)
		End If
		' compare to get maxPercentChange
		If resultArray(i, percentChange) > maxPercentChangeValue Then
			maxPercentChangeValue = resultArray(i, percentChange)
			maxPercentChangeName = resultArray(i, ticker)
		End If
		' compare to get maxTotalVol
		If resultArray(i, totalVol) > maxTotalVolValue Then
			maxTotalVolValue = resultArray(i, totalVol)
			maxTotalVolName = resultArray(i, ticker)
		End If
        Next i
	' --- output max and min variables
	' ticker names to col 16
	Cells(2, 16) = maxPercentChangeName
	Cells(3, 16) = minPercentChangeName
	Cells(4, 16) = maxTotalVolName
	' values to col 17
	Cells(2, 17) = maxPercentChangeValue
	Cells(3, 17) = minPercentChangeValue
	Cells(4, 17) = maxTotalVolValue

	' --- data labels
	

	' turn screen updating back on
	Application.ScreenUpdating = True
End Sub

         

        




' Sub that formats a column based on positive or negative value

'



