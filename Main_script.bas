Sub Ticker()

'Define variables as well as looping through worksheets.
Dim ws As Worksheet
Dim Endrow As Long

'looping through worksheets in book
For Each ws In Worksheets

'Adding headers to top of worksheet

ws.Cells(1, "I").Value = "Ticker"
ws.Cells(1, "J").Value = "Yearly Change"
ws.Cells(1, "K").Value = "Percent Change"
ws.Cells(1, "L").Value = "Total Stock Volume"

'Defining the variables

Dim Ticker As String
Dim Volume As Double
Dim Ticker_amt As Integer
Dim Open_price As Double
Dim Close_price As Double
Dim Yearly_change As Double
Dim Percent_change As Double

'Reset the ticker amount
Ticker_amt = 0

'assigning the last row in column A
Endrow = ws.Cells(Rows.Count, "A").End(xlUp).Row


'The following will run through each row of the worksheets
For Row = 2 To Endrow
    Ticker = ws.Cells(Row, 1).Value
    
    If Open_price = 0 Then
        Open_price = ws.Cells(Row, 3).Value
    End If
    
Volume = Volume + ws.Cells(Row, 7).Value

'Checking if the ticker name is different
If ws.Cells(Row + 1, 1).Value <> Ticker Then
    
    Ticker_amt = Ticker_amt + 1
    ws.Cells(Ticker_amt + 1, 9) = Ticker
    
    Close_price = ws.Cells(Row, 6).Value
    
    Yearly_change = Close_price - Open_price
    
    ws.Cells(Ticker_amt + 1, 10).Value = Yearly_change
    
    If Yearly_change >= 0 Then
        ws.Cells(Ticker_amt + 1, 10).Interior.ColorIndex = 4
        
    Else
        ws.Cells(Ticker_amt + 1, 10).Interior.ColorIndex = 3
    
    End If
    
'Calculating the percent changes
If Open_price = 0 Then
    Percent_change = 0
    
Else
    Percent_change = (Yearly_change / Open_price)

End If

    ws.Cells(Ticker_amt + 1, 11).Value = Format(Percent_change, "Percent")
    
'Calculating the total stock volume
    ws.Cells(Ticker_amt + 1, 12).Value = Volume
    
'Resetting all the values
Volume = 0
Yearly_change = 0
Ticker = ""
Open_price = 0
Percent_change = 0
Close_price = 0

End If

Next Row
Next ws

End Sub
