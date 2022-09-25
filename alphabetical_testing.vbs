Attribute VB_Name = "Module1"
Sub stock()

'Declare Worksheet and other variables
Dim year_change As Double
Dim open_value As Double
Dim close_value As Double
Dim ticker As String
Dim counter As Integer
Dim stock_volume As Double
Dim percent_change As Double
Dim ws As Worksheet

'Looping through all the worksheets
For Each ws In Worksheets

'Count the number of number of rows in the first column
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Setting header for the values
 ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total stock Volume")

'Setting counter and open_value
 counter = 1
 open_value = 0

'Looping through the rows by the ticker names
 For i = 2 To lastrow

'Searches for when the value of the next cell is different from the current cell
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Increasing the counter when the condition meets and passing the corresponding close_value and ticker name
    counter = counter + 1
    close_value = ws.Cells(i, 6).Value
    ticker = ws.Cells(i, 1).Value
   
'passing the ticker names to the excel
    ws.Cells(counter, 9).Value = ticker
   
'Calculating the yearly change and passing the value to excel
    year_change = close_value - open_value
    'year_change = Application.WorksheetFunction.Round(year_change, 2)
    ws.Cells(counter, 10).Value = year_change
   
'Calculating the percent change and passing them to the excel
    percent_change = (year_change / open_value) * 100
    ws.Cells(counter, 11).Value = percent_change & "%"
   
'Calculating stock volume
    ws.Cells(counter, 12).Value = stock_volume + ws.Cells(i, 7).Value

'Setting color code for the yearly change
 If ws.Cells(counter, 10).Value < 0 Then
    ws.Cells(counter, 10).Interior.ColorIndex = 3
 
 Else
    ws.Cells(counter, 10).Interior.ColorIndex = 4
   
 End If

'clearing the values
   stock_volume = 0
   open_value = 0

 Else
'calculating the stock volume
   stock_volume = stock_volume + ws.Cells(i, 7).Value
   
'fetching the open_value
 If open_value = 0 Then
    open_value = ws.Cells(i, 3).Value
 End If

 End If

 Next i


'Label the Summery table headers
 ws.Cells(3, 15).Value = "Greatest % increase"
 ws.Cells(4, 15).Value = "Greatest % decrease"
 ws.Cells(5, 15).Value = "Greatest total volume"
 ws.Cells(2, 16).Value = "Ticker"
 ws.Cells(2, 17).Value = "Value"

'Finding the last row of the table
 lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Looping through the summary table
 For i = 2 To lastrow_summary_table

'Finding the Max percent change and the corresponding ticker name
  If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
     ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
     ws.Cells(3, 17).Value = ws.Cells(i, 11).Value

'Finding the Min percent change and the corresponding ticker name
 ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(i, 11).Value

'Finding the Max Volume and its ticker name
 ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
    ws.Cells(5, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(5, 17).Value = ws.Cells(i, 12).Value

 End If

 Next i

Next ws
End Sub




