Sub first()

    'Set initial variable for ticker
    Dim Ticker_Name As String

    'Set values for starting stock and ending stock
    Dim Open_Stock As Double
    Open_Stock = 0
    Dim End_Stock As Double
    End_Stock = 0
    Dim Yearly_Change As Double
    Yearly_Change = 0
    Dim Total_Stock As Double
    Total_Stock = 0
    
    ' Keep track of the location for each Ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Define Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    'Obtain last row
    Dim LR As Long
    LR = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row

    'Loop Through all ticker
    For i = 2 To LR
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            'Set Ticker Name
            Ticker_Name = Cells(i, 1).Value

            Open_Stock = Open_Stock + Cells(i, 3).Value
            End_Stock = End_Stock + Cells(i, 6).Value
            Yearly_Change = Open_Stock - End_Stock
            Total_Stock = Total_Stock + Cells(i, 7).Value
            
            'Place in ticker column
            Cells(Summary_Table_Row, 9).Value = Ticker_Name

            'Place the Yearly change
            Cells(Summary_Table_Row, 10).Value = Yearly_Change

            'Do % change
            Cells(Summary_Table_Row, 11).Value = Yearly_Change / Open_Stock

            'Total Stock Volume
            Cells(Summary_Table_Row, 12).Value = Total_Stock

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1



            'Reset Total
            Open_Stock = 0
            End_Stock = 0
            Yearly_Change = 0
            Total_Stock = 0

        Else

            ' Add to the Brand Total
            Open_Stock = Open_Stock + Cells(i, 3).Value
            End_Stock = End_Stock + Cells(i, 6).Value
            Yearly_Change = Open_Stock - End_Stock
            Total_Stock = Total_Stock + Cells(i, 7).Value


        End If

'Conditional for positive or negative change
If Cells(i, 10) > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
ElseIf Cells(i, 10) < 0 Then
    Cells(i, 10).Interior.ColorIndex = 3
End If

'Percent
Cells(i, 11).NumberFormat = "0.00%"
    Next i
    

End Sub