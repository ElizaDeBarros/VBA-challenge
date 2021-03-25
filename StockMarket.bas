Attribute VB_Name = "Module1"
Sub StockMarket()

'Hides screen updates while the code is being executed
Application.ScreenUpdating = False

'Finds the last non empty row of column A of the spreadsheet
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Volume = Cells(2, 7)
IndexRow = 2
B = 2
COUNTER = 2

'Writes heads for the out put data
Cells(1, 9) = "Ticker"
Cells(1, 10) = "YearlyChange"
Cells(1, 11) = "PercentChange"
Cells(1, 12) = "Total Stock Volume"
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 15) = "Greatest % Increase"
Cells(3, 15) = "Greatest % Decrease"
Cells(4, 15) = "Greatest Total Volume"


'loops through the second row (row that contains data) to the last row with data
For i = 2 To LastRow
'Increments the row number as the loop is being executed
IndexRow = IndexRow + 1
    'conditional that looks for all rows with the same ticker and calculate the statistics
    If Cells(i, 1) = Cells(i + 1, 1) Then
        Volume = Volume + Cells(i + 1, 7)
        YC = Cells(IndexRow, 6) - Cells(B, 3)
            If Cells(B, 3) <> 0 Then
            PC = ((Cells(IndexRow, 6) / Cells(B, 3)) - 1)
            Else
            PC = 0
            End If
    'if above is false, it increments the output row for the next ticker
    Else
    B = IndexRow
    COUNTER = COUNTER + 1
    Volume = Cells(IndexRow, 7)
    End If
'writes results on columns I to L
Cells(COUNTER, 9) = Cells(i, 1)
Cells(COUNTER, 10) = YC
Cells(COUNTER, 11) = PC
Cells(COUNTER, 12) = Volume
Next

'Formats output data
LastRow = Cells(Rows.Count, 9).End(xlUp).Row
Range(Cells(2, 10), Cells(LastRow, 10)).NumberFormat = "0.00"
Range(Cells(2, 11), Cells(LastRow, 11)).NumberFormat = "0.00%"

'finds last non empty row of column I
LastRow = Cells(Rows.Count, 9).End(xlUp).Row

'Formats the colors of column J based on conditionals
For i = 2 To LastRow
    If Cells(i, 10) < 0 Then
    Range(Cells(i, 10), Cells(i, 10)).Interior.Color = RGB(255, 0, 0)
    Else
    Range(Cells(i, 10), Cells(i, 10)).Interior.Color = RGB(102, 255, 51)
    End If
Next
    
    
'**********************************************BONUS*************************************************
'Once the output is completely written, if look for maximum and minimum value of the range in column K and the maximum value of column L
RGPC = Range(Cells(2, 11), Cells(LastRow, 11))
RGVol = Range(Cells(2, 12), Cells(LastRow, 12))
MaxInc = Application.Max(RGPC)
MaxDec = Application.Min(RGPC)
MaxVol = Application.Max(RGVol)

'writes and formats results on column Q
Cells(2, 17) = MaxInc
Range(Cells(2, 17), Cells(2, 17)).NumberFormat = "0.00%"
Cells(3, 17) = MaxDec
Range(Cells(3, 17), Cells(3, 17)).NumberFormat = "0.00%"
Cells(4, 17) = MaxVol

'Search the row number that corresponds to each value on column Q and writes the respective thicker of column I on columns P
RowNum = Application.WorksheetFunction.Match(MaxInc, Range(Cells(2, 11), Cells(LastRow, 11)), 0)
Cells(2, 16) = Cells(RowNum + 1, 9)
RowNum = Application.WorksheetFunction.Match(MaxDec, Range(Cells(2, 11), Cells(LastRow, 11)), 0)
Cells(3, 16) = Cells(RowNum + 1, 9)
RowNum = Application.WorksheetFunction.Match(MaxVol, Range(Cells(2, 12), Cells(LastRow, 12)), 0)
Cells(4, 16) = Cells(RowNum + 1, 9)

'Shows updated screen
Application.ScreenUpdating = True

MsgBox "Done!"
End Sub

