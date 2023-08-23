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
        ' there are no blanks in the <ticker> column, a blank will be taken as the end of the data for that sheet
' ---
' CalculateYearData: 
' Sub procedure that loops through one worksheet that contains daily data for a year and outputs yearly summary data
' Reads in data for each ticker, determines its earliest and latest dates and uses that to get first_open (the first opening price) and last_close (the last closing price)
' first_open and last_close are then used to calculate the yearly change ($) and percent change.
' Also, sums all vol values for a ticker to get total_vol
' Output colums that will be placed starting on column 9 of the same worksheet: 
        ' ticker
                ' ticker symbol, grouping variable key
        ' yearly change ($)
                ' yearly change from the opening price at the beginning of a year to the closing price at the end of that year.
                ' year_change = first_open - last_close
        ' percent change
                ' percentage change from the opening price at the beginning of a year to the closing price at the end of that year.
                ' percent_change = year_change / first_open
        ' total stock volume
                ' total stock volume of the stock, sum of volume of stock for all dates in a year.
                ' total_vol = vol_1 + vol_2 + ... + vol_n
Sub CalculateYearData():
' create new dictionary object to store data that is read in
' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object
' keys will be the ticker names
' items will be arrays with 5 elements:
        ' 0     first_date
        ' 1     last_date
        ' 2     first_open
        ' 3     last_open
        ' 4     total_vol
Dim stock_dict
Set stock_dict = CreateObject("Scripting.Dictionary")
' Dim stock_dict As Scripting.dictionary
' Set stock_dict = New Scripting.dictionary 

' this_row counter to loop through rows
Dim this_row as Integer
this_row = 2 ' starts at 2 because of the header row
' do this until ticker col is empty
do until Cells(this_row, 1).Value = ""
        ' read row ticker value as this_ticker
        this_ticker = Cells(this_row, 1)
        ' read this_date
        this_date = Cells(this_row, 2)
        ' if this_ticker exists in dictionary as a key,
        If stock_dict.exists(this_ticker) then
                ' using data for this_ticker
                With stock_dict.Item(this_ticker)
                        ' if this_date is lower than first_date, 
                        If this_date < .Item(0) then
                                ' replace first_date with this_date
                                .Item(0) = this_date
                                ' replace first_open with this_open
                                .Item(2) = Cells(this_row, 3)
                        End If
                        ' if this_date is higher than last_date,
                        If this_date < .Item(1) then
                                ' replace last_date with this_date
                                .Item(1) = this_date
                                ' replace last_close with this_close
                                .Item(3) = Cells(this_row, 6)
                        End If
                        ' and add this_vol to total_vol
                        .Item(4) = .Item(4) + Cells(this_row, 7)
                End With
        ' else, key doesn't exist so add new key as this_ticker
        else
                stock_dict.Add(this_ticker, Array(_
                        ' first_date is this_date
                        this_date,_
                        ' last_date is this_date
                        this_date,_
                        ' first_open is this_open
                        Cells(this_row, 3),_
                        ' last_close is this_close
                        Cells(this_row, 6),_
                        ' total_vol is this_vol
                        Cells(this_row, 7)_
                        )_
                )
        End If
        ' increase this_row counter
        this_row = this_row + 1
loop

' output columns
' output starts with header row at column 9

' then puts down data from the dictionary
For output_row = 2 to stock_dict.Count + 1
        ' loop through dictionary keys: https://stackoverflow.com/questions/1296225/iterate-over-vba-dictionaries
        For each ticker In stock_dict.Keys()
                Cells(output_row, 9) = ticker
                ' using current ticker
                With stock_dict(ticker)
                ' year_change = first_open - last_close
                Cells(output_row, 10) = .Item(2) - .Item(3)
                ' percent_change = year_change / first_open
                Cells(output_row, 11) = Cells(output_row, 10) / .Item(2)
                ' total_vol
                Cells(output_row, 11) = .Item(4)
                End With
        Next ticker
Next output_row

End Sub

' Sub that formats a column based on positive or negative value

' 
