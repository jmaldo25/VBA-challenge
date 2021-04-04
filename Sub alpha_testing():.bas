' VBA Homework
' Solved to seperate tickers
' Solved for yearly change from opening price to closing price at end of that year
' Solved percentage change from opening price at beginning of year to closing price at end of year
' Solved total volume of stock

Sub alpha_testing():

Dim ticker As String
Dim row As Integer
Dim change As Double
Dim precent_change As Double
Dim volume As Double
Dim open_t As Double
Dim close_t As Double
row = 0

Dim ticker_summary_row As Integer
ticker_summary_row = 2

open_t = Cells(2, 3).Value

For i = 2 To 70926

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    ticker = Cells(i, 1).Value

    close_t = Cells(i, 6).Value

    volume = volume + Cells(i, 7).Value

    change = close_t - open_t

    percent_change = change / open_t

    Range("I" & ticker_summary_row).Value = ticker

    Range("L" & ticker_summary_row).Value = volume

    Range("J" & ticker_summary_row).Value = change

    Range("K" & ticker_summary_row).Value = percent_change

    ticker_summary_row = ticker_summary_row + 1

    volume = 0

    Else: volume = volume + Cells(i, 7).Value

    End If

Next i

End Sub


