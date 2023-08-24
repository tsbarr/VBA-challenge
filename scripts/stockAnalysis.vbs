' UofT SCS EdX Data Bootcamp: Challenge 2
' Script Author: Tania Barrera (tsbarr)
' ---
' Assumptions:
        ' each sheet contains data from a single year
        ' data from the year is organized as daily data, with a row containing data for a stock ticker in a date
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

Sub test():
        ' variable declaration for reading data in
        Dim lastRow As Long
        Dim inputArray() As Variant
	Dim resultArray() As Variant ' type variant so we can use different data types
	Dim uniqueTickers As New Collection
        Dim thisTicker As String
        Dim i As Long, j as Long ' for loop counters / index within array iterators
	
        
        ' Find last row with data in <ticker> column (i.e. 1)
        ' Source: https://www.excelcampus.com/vba/find-last-row-column-cell/
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row

        ' use lastRow to read in range of input data, no headers
        inputArray = Range("A2:G" & lastRow)

        ' get unique ticker values
        ' Source: https://stackoverflow.com/a/3017973
        On Error Resume Next ' if ticker is already in collection, skip, since it will throw error
	' note: inputArray is 1-based because it comes from a Range object
        For i = 1 To UBound(inputArray, 1) 
        thisTicker = inputArray(i, 1)
        uniqueTickers.Add thisTicker, thisTicker
        Next
        On Error GoTo 0

        ' use number of uniqueTickers to declare resultArray, since we now know the number of rows we need
         ' we need an array and not collection so we can perform operations, as per: https://excelmacromastery.com/excel-vba-collections/
        ' see here for declaring dynamic arrays:
        ' https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/declaring-arrays
        
        ReDim resultArray(uniqueTickers.Count - 1, 7) ' array dimensions, note: arrays are 0-based, but collections are 1-based
        ' resultArray has 8 columns:
                ' 0     ticker
                ' 1     first_date
                ' 2     first_open
                ' 3     last_date
                ' 4     last_close
                ' 5     year_change = first_open - last_open
                ' 6     percent_change = year_change / first_open
                ' 7     total_vol

        ' write uniqueTickers to resultArray first column (0)    
        For i = 1 To uniqueTickers.Count
        resultArray(i - 1, 0) = uniqueTickers(i) ' note: resultArray is 0-based, but collections are 1-based
        Next i

	' iterate within inputArray to obtain first and last date, first open, last close and sum of vol values
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
        		If resultArray(j, 0) = thisTicker Then 'If value found
            			' current row j is index in resultArray of thisTicker
            			Exit For
        		End If
    		Next j
		' check if this is the first time we find thisTicker by checking if value first_date is empty
		If IsEmpty(resultArray(j, 1)) Then 
			' initialize with current row info
			resultArray(j, 1) = thisDate 	' first_date
			resultArray(j, 2) = thisOpen	' first_open
			resultArray(j, 3) = thisDate 	' last_date
			resultArray(j, 4) = thisClose	' last_close
			resultArray(j, 7) = thisVol	' total_vol
		Else ' we have seen thisTicker before
			' check if this date is earlier than first_date
			If thisDate < resultArray(j, 1) Then
			' replace first_date with thisDate
			resultArray(j, 1) = thisDate
			' and first_open with thisOpen
			resultArray(j, 2) = thisOpen
			End If
			' check if thisDate is later than last_date
			If thisDate > resultArray(j, 3) Then
			' replace last_date with thisDate
			resultArray(j, 3) = thisDate
			' and last_close with thisClose
			resultArray(j, 4) = thisClose
			End If
			' add thisVol to total_vol
			resultArray(j, 7) = resultArray(j, 7) + thisVol
		End If
	Next i
	' iterate within resultArray to calculate year and percent change
	


        ' write result array to worksheet
        For i = 0 To UBound(resultArray, 1)
        Cells(i + 2, 9) = resultArray(i, 0)	' column ticker to column I
        Cells(i + 2, 10) = resultArray(i, 4)	' column year_change to column J
	Cells(i + 2, 11) = resultArray(i, 5)	' column percent_change to column K
        Cells(i + 2, 12) = resultArray(i, 6)	' column total_vol to column L
        Next i
End Sub

         

        




' Sub that formats a column based on positive or negative value

'


