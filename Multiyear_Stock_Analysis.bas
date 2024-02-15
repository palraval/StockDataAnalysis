Attribute VB_Name = "Module1"
Sub Unique_Ticker(ws As Worksheet)

Dim final_row, Ticker, i As Long
final_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

Ticker = 2
For i = 2 To final_row
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
        Ticker = Ticker + 1
    End If
Next i

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 9).Characters(Start:=1, Length:=6).Font.Bold = True


End Sub

Sub Yearly_change(ws As Worksheet)


Dim store_open, store_close, co, i As Long
Dim change, percent_change As Double

final_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

co = 2
store_open = 0
store_close = 0

For i = 2 To final_row
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        store_open = ws.Cells(i, 3).Value
    ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        store_close = ws.Cells(i, 6).Value
    End If
    If (store_open <> 0 And store_close <> 0) Then
        change = store_close - store_open
        percent_change = (change / store_open) * 100
        ws.Cells(co, 10).Value = change
        ws.Cells(co, 11).Value = percent_change
        co = co + 1
        store_open = 0
            store_close = 0
    End If
Next i

ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 10).Characters(Start:=1, Length:=13).Font.Bold = True
ws.Cells(1, 11).Value = "Percent Change (%)"
ws.Cells(1, 11).Characters(Start:=1, Length:=18).Font.Bold = True


End Sub

Sub stock_volume(ws As Worksheet)

Dim final_row, Row, Sum, jum As Long

final_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

Sum = 0
jum = 2
For Row = 3 To final_row:
    If ws.Cells(Row, 1).Value = ws.Cells(Row - 1, 1).Value Then
        Sum = ws.Cells(Row - 1, 7) + Sum
    ElseIf ws.Cells(Row, 1).Value <> ws.Cells(Row - 1, 1).Value Then
        Sum = ws.Cells(Row - 1, 7) + Sum
        ws.Cells(jum, 12).Value = Sum
        Sum = 0
        jum = jum + 1
    End If
Next Row
     
ws.Cells(1, 12).Value = "Cummulative Stock Volume"
ws.Cells(1, 12).Characters(Start:=1, Length:=25).Font.Bold = True


End Sub

Sub yearly_conditional(ws As Worksheet)

Dim new_row, yearchange As Long


new_row = ws.Cells(Rows.Count, 10).End(xlUp).Row

For yearchange = 2 To new_row
    If ws.Cells(yearchange, 10) < 0 Then
        ws.Cells(yearchange, 10).Interior.ColorIndex = 3
    ElseIf ws.Cells(yearchange, 10) > 0 Then
        ws.Cells(yearchange, 10).Interior.ColorIndex = 10
    ElseIf ws.Cells(yearchange, 10) = 0 Then
        ws.Cells(yearchange, 10).Interior.ColorIndex = 44
    End If
Next yearchange

End Sub

Sub percent_conditional(ws As Worksheet)

Dim rows_again, perchange As Long

rows_again = ws.Cells(Rows.Count, 11).End(xlUp).Row

For perchange = 2 To rows_again
    If ws.Cells(perchange, 11) < 0 Then
        ws.Cells(perchange, 11).Interior.ColorIndex = 3
    ElseIf ws.Cells(perchange, 11) > 0 Then
        ws.Cells(perchange, 11).Interior.ColorIndex = 10
    ElseIf ws.Cells(perchange, 11) = 0 Then
        ws.Cells(perchange, 11).Interior.ColorIndex = 44
    End If
Next perchange


End Sub

Sub greatest_increase(ws As Worksheet)

Dim row11, highest_value, k, highest_k, lowest_value, lowest_k, highest_volume, highest_k_volume As Long


row11 = ws.Cells(Rows.Count, 11).End(xlUp).Row

highest_value = 0
lowest_value = 0
highest_volume = 0

For k = 2 To row11
    If ws.Cells(k, 11).Value > highest_value Then
        highest_value = ws.Cells(k, 11).Value
        highest_k = k
    ElseIf ws.Cells(k, 11).Value < lowest_value Then
        lowest_value = ws.Cells(k, 11).Value
        lowest_k = k
    End If
    If ws.Cells(k, 12).Value > highest_volume Then
        highest_volume = ws.Cells(k, 12).Value
        highest_k_volume = k
    End If
Next k


ws.Cells(3, 16).Value = "Ticker"
ws.Cells(3, 16).Characters(Start:=1, Length:=6).Font.Bold = True
ws.Cells(3, 17).Value = "Value"
ws.Cells(3, 17).Characters(Start:=1, Length:=5).Font.Bold = True
ws.Cells(4, 15).Value = "Greatest Percent Increase"
ws.Cells(4, 15).Characters(Start:=1, Length:=25).Font.Bold = True
ws.Cells(4, 16).Value = Cells(highest_k, 9).Value
ws.Cells(4, 17).Value = highest_value
ws.Cells(5, 15).Value = "Greatest Percent Decrease"
ws.Cells(5, 15).Characters(Start:=1, Length:=25).Font.Bold = True
ws.Cells(6, 15).Value = "Greatest Cummulative Volume"
ws.Cells(6, 15).Characters(Start:=1, Length:=27).Font.Bold = True
ws.Cells(5, 17).Value = lowest_value
ws.Cells(5, 16).Value = Cells(lowest_k, 9).Value
ws.Cells(6, 17).Value = highest_volume
ws.Cells(6, 16).Value = Cells(highest_k_volume, 9).Value

ws.Range("O3:Q6").Interior.ColorIndex = 34


End Sub

Sub WorksheetLoop()


Dim ws As Worksheet

For Each ws In ActiveWorkbook.Sheets
    Call Unique_Ticker(ws)
    Call Yearly_change(ws)
    Call stock_volume(ws)
    Call yearly_conditional(ws)
    Call percent_conditional(ws)
    Call greatest_increase(ws)
Next ws


End Sub


