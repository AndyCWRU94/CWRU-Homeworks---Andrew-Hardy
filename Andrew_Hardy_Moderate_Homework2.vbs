Sub Stockanalysis()

' Every worksheet in workbook
Dim ws As Worksheet
For Each ws In Worksheets

'headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Volume"

'variables
Dim ticker_symbol As String
Dim volume As Double
Dim tablerow As Double
Dim last As Double
Dim openprice As Double
Dim closeprice As Double
Dim totalchange As Double
Dim percentchange As Double

'initial variable values
tablerow = 2
volume = 0
openprice = 0
closeprice = 0
totalchange = 0
percentchange = 0

'set last variable to last row of data
last = Cells(Rows.Count, 1).End(xlUp).Row

'loop through all tickers in sheet
For i = 2 To last

        'if ticker changed, then
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        openprice = ws.Cells(i, 3).Value

        'retain openprice value
        Else
        openprice = openprice
        End If

    '
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

    'set ticker name
    ticker_symbol = ws.Cells(i, 1)

    ' add to total volume on every row
    volume = volume + ws.Cells(i, 7).Value

    ' add row closeprice to total closeprice
    closeprice = ws.Cells(i, 6).Value

    ' compute totalchange and percentchange
    totalchange = closeprice - openprice
    percentchange = totalchange / openprice

    ' add ticker name and volume to summary table
    ws.Range("I" & tablerow).Value = ticker_symbol
    ws.Range("J" & tablerow).Value = totalchange
    ws.Range("K" & tablerow).Value = percentchange
    ws.Range("L" & tablerow).Value = volume

    ' set up summary table for next change in ticker name
    tablerow = tablerow + 1

    ' clear total volume for next time
    volume = 0
    Else
    volume = volume + ws.Cells(i, 7).Value
   'add volume to summary table
    ws.Range("J" & tablerow).Value = volume

    End If

Next i

Next ws

End Sub