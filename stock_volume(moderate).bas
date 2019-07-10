Attribute VB_Name = "Module1"
Sub Stocks()

'Loop through all sheets
   Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set an initial variable for holding the ticket value
Dim Ticker As String
' Set an initial variable for holding the total stock value per ticker
Dim Total_Stock As Double
Total_Stock = 0
'keep track of ticker values
Dim Ticker_Value_Row As Integer
Ticker_Value_Row = 2
'name the headers
Cells(1, "I").Value = "Ticker"
Cells(1, "L").Value = "Total Stock"
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Change"
'add variables
Dim open_value As Double
Dim close_value As Double
Dim yearly_change As Double
Dim percent_change As Double


'Loop through all ticker values
For i = 2 To LastRow
'Check if we are within the same ticker value, if it is not
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
'Set the ticker
Ticker = Cells(i, 1).Value
'Add to Total Stock
Total_Stock = Total_Stock + Cells(i, 7).Value
'set open value
open_value = Cells(2, 3).Value
'set close value
close_value = Cells(i, 6).Value
'Calculate Yearly Change
yearly_change = close_value - open_value
'Calculate percent change
If open_value = 0 And close_value = 0 Then
    percent_change = 0
    ElseIf open_value = 0 And close_value <> 0 Then
    percent_change = 1
    Else: percent_change = yearly_change / open_value
'Print percent change
Range("K" & Ticker_Value_Row).Value = percent_change
Range("K" & Ticker_Value_Row).NumberFormat = "0.00%"
End If

'Print yearly change
Range("J" & Ticker_Value_Row).Value = yearly_change

'Print Ticker in the Ticker Value Row
Range("I" & Ticker_Value_Row).Value = Ticker
'Print Stock Amount to the Ticker Value Row
Range("L" & Ticker_Value_Row).Value = Total_Stock
'Add one to Ticker_Value_Row
Ticker_Value_Row = Ticker_Value_Row + 1
'Reset Total_Stock
Total_Stock = 0
'If the cell imediately following a row is the same ticker
Else
'Add to the Total Stock
Total_Stock = Total_Stock + Cells(i, 7).Value
End If
'Add color to yearly change
If Cells(i, 10).Value > 0 Or Cells(i, 10).Value = 0 Then
Cells(i, 10).Interior.ColorIndex = 10
ElseIf Cells(i, 10).Value < 0 Then
Cells(i, 10).Interior.ColorIndex = 3
End If

Next i
Next ws

End Sub

