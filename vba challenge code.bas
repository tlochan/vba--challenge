Attribute VB_Name = "Module1"
Sub allws():
 Dim ws As Worksheet
 For Each ws In ThisWorkbook.Worksheets
  ws.Activate
 stock_summary_table
 Next ws
 End Sub
 
Sub stock_summary_table():

'Summary Table Headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"


'lastrow function for column A
Dim lastrow As Long
lastrow = WorksheetFunction.CountA(Columns("A:A"))

'variable total for total stock volume
Dim total As Double
total = 0

'variable for a row for each unique ticker in the summary table
Dim tablerow As Integer
tablerow = 2

'conditional for loop -> if ticker is same as next row
For i = 2 To lastrow
    
'variables for the summary table
    Dim ticker As String
    ticker = Cells(i, 1).Value
    Dim initial As Double
    Dim final As Double
    Dim yearly_change As Double
    Dim yearly_percent As Double
    
'conditional to obtain the first opening value
     If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
     initial = Cells(i, 3).Value
     
     End If

'conditional counting total stock volume
    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    Dim vol As Double
    vol = Cells(i, 7).Value
    total = total + vol

    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'print ticker name in the summary table
    Cells(tablerow, 9).Value = ticker
    
'grab the closing as final stock price
    final = Cells(i, 6).Value
    
'yearly change calc and print
    yearly_change = final - initial
    Cells(tablerow, 10).Value = yearly_change
    
'conditional formatting for yearly change
    If yearly_change > 0 Then
    Cells(tablerow, 10).Interior.ColorIndex = 4
    
    ElseIf yearly_change < 0 Then
    Cells(tablerow, 10).Interior.ColorIndex = 3
    
    Else: Cells(tablerow, 10).Interior.ColorIndex = 0
    
    End If
    
'conditional formatting for percent change
yearly_percent = (yearly_change / initial)

If yearly_percent > 0 Then
    Cells(tablerow, 11).Interior.ColorIndex = 4
    
    ElseIf yearly_percent < 0 Then
    Cells(tablerow, 11).Interior.ColorIndex = 3
    
    Else: Cells(tablerow, 11).Interior.ColorIndex = 0
    
    End If
    
'percent change calc and print
    yearly_percent = WorksheetFunction.Round((yearly_change / initial), 2)
    'Cells(tablerow, 11).Value = Str(yearly_percent) + " %"
    Cells(tablerow, 11).Value = yearly_percent

    
'add last volume, print, then reset to zero
    total = total + Cells(i, 7).Value
    Cells(tablerow, 12).Value = total
    total = 0
'increase summary row by 1 for next ticker
    tablerow = tablerow + 1

    End If
    
    Next i
    
'minmax table code

'Headers for second table
Range("n2").Value = "Greatest % Increase"
Range("n3").Value = "Greatest % Decrease"
Range("n4").Value = "Greatest Total Volume"
Range("o1").Value = "Ticker"
Range("P1").Value = "Value"

'variables for values and associated tickers
Dim max As Double
Dim min As Double
Dim max_vol As Double

Dim maxticker As String
Dim minticker As String
Dim volticker As String

'min/max functions
max = WorksheetFunction.max(Range("K:K"))
min = WorksheetFunction.min(Range("K:K"))
max_vol = WorksheetFunction.max(Range("L:L"))

'printing values in table
Range("p2").Value = max
Range("p3").Value = min
Range("p4").Value = max_vol

'searching first summary table for tickers using for loop
Dim lastrow2 As Long
lastrow2 = WorksheetFunction.CountA(Columns("I:I"))

For i = 2 To lastrow2
If Cells(i, 11).Value = max Then
maxticker = Cells(i, 9).Value
End If

If Cells(i, 11).Value = min Then
minticker = Cells(i, 9).Value
End If

If Cells(i, 12).Value = max_vol Then
volticker = Cells(i, 9).Value
End If

Next i

'printing tickers
Range("o2") = maxticker
Range("O3") = minticker
Range("O4") = volticker

End Sub


