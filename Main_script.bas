Attribute VB_Name = "Module1"
Sub Ticker()

'Define variables as well as looping through worksheets.
Dim ws As Worksheet
Dim endrow As Long

'looping through worksheets in book
For Each ws In Worksheets

'Adding headers to top of worksheet

ws.Cells(1, "I").Value = "Ticker"
ws.Cells(1, "J").Value = "Yearly Change"
ws.Cells(1, "K").Value = "Percent Change"
ws.Cells(1, "L").Value = "Total Stock Volume"

'Defining the variables

Dim Ticker As String
Dim volume As Double





Next ws

End Sub
