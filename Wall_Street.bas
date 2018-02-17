Attribute VB_Name = "Wall_Street"
Sub Wall_Street():

'Create a script that will loop through each year of stock data and grab
'the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the
'total volume.
'Your result should look as follows.

'1: Assign Variables
    Dim stock_year As Long
    Dim total_volume As Double
    Dim ticker_symbol As String
    Dim ticker_table_row As Integer
    Dim percent_change As Double
    Dim price_change As Double
    Dim percent_increase As Double
    Dim greatest_pct_increase As Double
    Dim greatest_pct_decrease As Double
    Dim greatest_volume As Double
    
    Dim greatest_pct_increase_ticker As String
    Dim greatest_pct_decrease_ticker As String
    Dim greatest_volume_ticker As String
    
    
    
    
'2: Values
        'Range(“B2”).Value = Format(Range("B2").Value, (“dd / mm / yyyy”))
        total_volume = 0
        ticker_table_row = 1
        Cells(1, 12) = "Total Stock Volume"
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Price Change"
        Cells(1, 11) = "Percent Change"
        Cells(2, 15) = "Greatest % Increase"
        Cells(3, 15) = "Greatest % Decrease"
        Cells(4, 15) = "Greatest Total Volume"
        Cells(1, 16) = "Ticker"
        Cells(1, 17) = "Value"
        year_open_price = Cells(2, 3).Value
        
'3: Loops
        'Change this to last row calculator
        For Row = 2 To 70926
                'stock_year = Cells(Row, 2).Value
                ticker_symbol = Cells(Row, 1).Value
                'Col = Cells(Row, 6).Value
                price_change = 0
                percent_change = 0
'Add to total
                total_volume = total_volume + Cells(Row, 7).Value
'4: Conditionals
            If year_open_price = 0 Then
                year_open_price = Cells(Row, 3)
            End If
            If Cells(Row + 1, 1) <> ticker_symbol Then
                ticker_table_row = ticker_table_row + 1
                ticker_symbol = Cells(Row, 1).Value
'For Moderate: Percent Change open/close
                'First open non-0 price
                
'For Moderate: Price Change open/close
                'First open non-0 price per ticker
                'Year close is last close per ticker
                year_close_price = Cells(Row, 6).Value
                price_change = year_close_price - year_open_price
                If year_open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = price_change / year_open_price
                End If
'For Hard: Greatest % Increase
                
'Greatest % Decrease



'Greatest Total Volume
                'greatest_volume = Cells(4, 17)
                'Cells(4, 17) = Application.WorksheetFunction.Max(Cells(Row, 7)).Value
                'Cells(4, 17) = Max(total_volume)
'5: Print Ticker in Table
                    Range("I" & ticker_table_row).Value = ticker_symbol
                    Range("J" & ticker_table_row).Value = price_change
                If (price_change < 0) Then
                    Range("J" & ticker_table_row).Interior.Color = RGB(255, 0, 0)
                Else:
                    Range("J" & ticker_table_row).Interior.Color = RGB(0, 255, 0)
                End If
'Print Percent_change in Table
                    Range("K" & ticker_table_row).Value = percent_change
                    Range("K" & ticker_table_row).Style = "Percent"
'6: Print Volume in Table
                    Range("L" & ticker_table_row).Value = total_volume

'7: Hard(greatest % increase)
    
                    
                    year_open_price = Cells(Row + 1, 3)
                If (percent_change > greatest_pct_increase) Then
                    greatest_pct_increase = percent_change
                    greatest_pct_increase_ticker = ticker_symbol
                    Cells(2, 16).Value = greatest_pct_increase_ticker
                    Cells(2, 17).Value = greatest_pct_increase
                    Cells(2, 17).NumberFormat = "0.00%"
                End If
'greatest % decrease
                 If (percent_change < greatest_pct_decrease) Then
                    greatest_pct_decrease = percent_change
                    greatest_pct_decrease_ticker = ticker_symbol
                    Cells(3, 16).Value = greatest_pct_decrease_ticker
                    Cells(3, 17).Value = greatest_pct_decrease
                    Cells(3, 17).NumberFormat = "0.00%"
                End If
'greatest volume
                 If (total_volume > greatest_volume) Then
                    greatest_volume = total_volume
                    greatest_volume_ticker = ticker_symbol
                    Cells(4, 16).Value = greatest_volume_ticker
                    Cells(4, 17).Value = greatest_volume
                End If
                    total_volume = 0
            End If
             
            
            

            




    Next Row





End Sub
