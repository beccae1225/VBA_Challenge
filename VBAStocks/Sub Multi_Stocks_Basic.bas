Sub Multi_Stocks()

For Each ws In Worksheets

'Cell Titles
ws.Cells(1, 10).Value = "Ticker Symbol"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"

'Declare variables
Dim Ticker_Symbol As String
Ticker_Symbol = ""

Dim Stock_Volume As Variant
Stock_Volume = 0

Dim Open_Price As Double
Open_Price = ws.Cells(2, 3).Value

Dim Close_Price As Double
Close_Price = 0

Dim Percent_Change As Double
Percent_Change = 0

Dim Yearly_Change As Double
Yearly_Change = 0

Dim Summary_Table_Row As Long
Summary_Table_Row = 2

Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To LastRow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        Ticker_Symbol = Cells(i, 1).Value
        Stock_Volume = Stock_Volume + Cells(i, 7).Value
        
        ws.Range("J" & Summary_Table_Row).Value = Ticker_Symbol
        ws.Range("M" & Summary_Table_Row).Value = Stock_Volume
        
        Close_Price = ws.Cells(i, 6).Value
        Yearly_Change = Close_Price - Open_Price

        ws.Range("K" & Summary_Table_Row).Value = Yearly_Change

            If (Open_Price = 0 And Yearly_Change = 0) Then
                Percent_Change = 0
            ElseIf (Yearly_Change <> 0 And Open_Price = 0) Then
                Percent_Change = 1
            Else: Percent_Change = Yearly_Change / Open_Price
                ws.Range("L" & Summary_Table_Row).Value = Percent_Change
                ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"

            End If
    
    Summary_Table_Row = Summary_Table_Row + 1

    Stock_Volume = 0
    Percent_Change = 0
    Yearly_Change = 0
    Open_Price = ws.Cells(i + 1, 3).Value

Else

    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

End If

Next i

'Color Format 

Dim color As Long
ColorRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

For j = 2 To ColorRow

    If ws.Cells(j, 11).Value <= 0 Then
        ws.Cells(j, 11).Interior.ColorIndex = 3

    Else
        ws.Cells(j, 11).Interior.ColorIndex = 4

    End If

Next j

Next ws

End Sub

