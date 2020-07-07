Attribute VB_Name = "Module1"
Sub stocks()


'set column headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Dim ticker As String
Dim rowcount As Variant
Dim open_amt, yearly_change, close_amt, pct_change As Double
    open_amt = 0
    close_amt = 0
    pct_change = 0
Dim volume As Variant
    volume = 0
'location of data
    Dim summary_row As Double
    summary_row = 2
'set varible for last row
    rowcount = Cells(Rows.Count, 1).End(xlUp).Row
' Iterate over dataset
    For r = 2 To rowcount
    
    
' check if same ticker if not the
    If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then

'set the stock ticker
    ticker = Cells(r, 1).Value
 'add to volume
    volume = volume + Cells(r, 7).Value
 'print ticker in summay
    Range("I" & summary_row).Value = ticker
'print volume in table
    Range("L" & summary_row).Value = volume
 'set open amount
    open_amt = Cells(r, 3).Value
'set close amount
    close_amt = Cells(r, 6).Value
'print yearly change
    yearly_change = close_amt - open_amt
    Range("J" & summary_row).Value = yearly_change
'print percentage change
    pct_change = yearly_change / open_amt * 100
    Range("K" & summary_row).Value = pct_change
    Range("K" & summary_row).NumberFormat = "0.00%"
'add one to summary table
     summary_row = summary_row + 1
 
'reset volume
    volume = 0
    
    Else
    
'add to volume total
volume = volume + Cells(r, 7).Value

    End If
    
    Next r
    
    
   End Sub
