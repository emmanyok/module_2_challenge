Attribute VB_Name = "Module1"
Sub module_challenge2()


'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets


'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker variable
Dim Ticker As String
Dim PriceChange As Double
Ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0
Dim total_stock_volume As Double


'Set initial and last row for worksheet
Dim Lastrow As Long
Dim i As Long
Dim j As Integer

'Define Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set new variables for prices and percent changes
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0
Dim price_change_percent As Double
price_change_percent = 0
total_stock_volume = 0


Dim TickerRow As Long: TickerRow = 1


'Do loop of current worksheet to Lastrow
For i = 2 To Lastrow

'Ticker symbol output
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
TickerRow = TickerRow + 1
Ticker = ws.Cells(i, 1).Value
ws.Cells(TickerRow, "I").Value = Ticker

'Calculate change in Price

close_price = ws.Cells(i, 6).Value
open_price = ws.Cells(i, 3).Value
price_change_percent = close_price - open_price
ws.Cells(TickerRow + 1, "J").Value = price_change_percent

'Calculate percentage change
percentage_change = (price_change_percent / open_price) * 100
ws.Cells(TickerRow, "K").Value = Format(percentage_change, "Percent")

'Add Total Stock
total_stock_volume = total_stock_volume + Cells(i, 7).Value
ws.Cells(TickerRow, "L").Value = total_stock_volume
total_stock_volume = 0

'Shade Ticker color
'If percent change value is greater than 0, shade cell green.
If price_change_percent > 0 Then
    Cells(TickerRow, "K").Interior.ColorIndex = 4
'If percent change value is less than 0, shade cell red.
ElseIf price_change_percent < 0 Then
    Cells(TickerRow, "K").Interior.ColorIndex = 3


End If

End If

Next i

Next ws

End Sub
