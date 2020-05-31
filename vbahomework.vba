Sub Stocks()
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim opening_price As Double
Dim closing_price As Double
Dim Summary_Table_Index As Integer
Dim greatest_percent_increase As Double
Dim greatest_percent_decrease As Double
Dim greatest_total_volume As Double
Dim GPI_ticker As String
Dim GPD_ticker As String
Dim GTV_ticker As String

Cells(1, 9).Value = "<Ticker>"
Cells(1, 10).Value = "<Yearly Change>"
Cells(1, 11).Value = "<Percent Change>"
Cells(1, 12).Value = "<Total Stock Volume>"
Cells(2, 14).Value = "Greatest Percent Increase"
Cells(3, 14).Value = "Greatest Percent Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

Summary_Table_Index = 2
total_stock_volume = 0
greatest_percent_increase = -9999
greatest_percent_decrease = 9999
greatest_total_volume = -9999

NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
For i = 2 To NumRows

opening_price = opening_price + Cells(i, 3).Value
closing_price = closing_price + Cells(i, 6).Value
total_stock_volume = total_stock_volume + Cells(i, 7).Value

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
ticker = Cells(i, 1).Value

Range("I" & Summary_Table_Index).Value = ticker
yearly_change = closing_price - opening_price
Range("J" & Summary_Table_Index).Value = yearly_change
percent_change = yearly_change / opening_price
Range("K" & Summary_Table_Index).Value = percent_change
Range("K" & Summary_Table_Index).NumberFormat = "0.00%"
Range("L" & Summary_Table_Index).Value = total_stock_volume

If percent_change > greatest_percent_increase Then
greatest_percent_increase = percent_change
GPI_ticker = Range("A" & Summary_Table_Index).Value

End If

If percent_change < greatest_percent_decrease Then
greatest_percent_decrease = percent_change
GPD_ticker = Range("A" & Summary_Table_Index).Value

End If

If total_stock_volume > greatest_total_volume Then
greatest_total_volume = total_stock_volume
GTV_ticker = Range("A" & Summary_Table_Index).Value

End If

If yearly_change >= 0 Then
Range("K" & Summary_Table_Index).Interior.ColorIndex = 4

End If

If yearly_change < 0 Then
Range("K" & Summary_Table_Index).Interior.ColorIndex = 3

End If

Summary_Table_Index = Summary_Table_Index + 1
yearly_change = 0
opening_price = 0
closing_price = 0
total_stock_volume = 0

End If

Next i

Cells(2, 15).Value = GPI_ticker
Cells(3, 15).Value = GPD_ticker
Cells(4, 15).Value = GTV_ticker
Cells(2, 16).Value = greatest_percent_increase
Cells(2, 16).NumberFormat = "0.00%"
Cells(3, 16).Value = greatest_percent_decrease
Cells(3, 16).NumberFormat = "0.00%"
Cells(4, 16).Value = greatest_total_volume

End Sub
