Sub stock_data()

For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly_change"
    ws.Cells(1, 11).Value = "Percent_change"
    ws.Cells(1, 12).Value = "Total_Stock Volume"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    
Dim ticker_name As String
Dim total_stock As Double
total_stock = 0
Dim Summary_table_row As Integer
Summary_table_row = 2

Dim yearly_change As Double
Dim percent_change As Double
yearly_change = 0
percent_change = 0


Dim open_price As Double
open_price = ws.Cells(2, 3).Value

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ticker_name = ws.Cells(i, 1).Value
ws.Range("I" & Summary_table_row).Value = ticker_name

yearly_change = ws.Cells(i, 6).Value - open_price
ws.Range("J" & Summary_table_row).Value = yearly_change

percent_change = yearly_change / open_price
ws.Range("K" & Summary_table_row).Value = percent_change
ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"



total_stock = ws.Cells(i, 7).Value + total_stock
ws.Range("L" & Summary_table_row).Value = total_stock

Summary_table_row = Summary_table_row + 1
total_stock = 0
open_price = ws.Cells(i + 1, 3).Value


Else
total_stock = total_stock + ws.Cells(i, 7).Value
End If
Next i

Dim summary_lastrow As Long
summary_lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

For i = 2 To summary_lastrow
If ws.Cells(i, 10).Value >= 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
Else
ws.Cells(i, 10).Interior.ColorIndex = 3
End If
Next i



Dim per_max As Double
per_max = 0
Dim per_min As Double
per_min = 0
Dim gst_vol As Double
gst_vol = 0


For i = 2 To summary_lastrow
If per_max < ws.Cells(i, 11).Value Then
per_max = ws.Cells(i, 11).Value
ws.Range("Q2").Value = per_max
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("P2").Value = ws.Cells(i, 9).Value

ElseIf per_min > ws.Cells(i, 11).Value Then
per_min = ws.Cells(i, 11).Value
ws.Range("Q3").Value = per_min
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("P3").Value = ws.Cells(i, 9).Value
End If
Next i

For i = 2 To summary_lastrow

If gst_vol < ws.Cells(i, 12).Value Then
gst_vol = ws.Cells(i, 12).Value
ws.Range("Q4").Value = gst_vol
ws.Range("P4").Value = ws.Cells(i, 9).Value
End If
Next i


Next ws


End Sub