Sub Stonks()
    ' This code was made possible with the use of Stack Overflow and Google. 
    ' We stand on the shoulders of giants!
    
'Loop through all Worksheets-------------------------------------

    For Each ws In Worksheets
    
'Headers and Formatting------------------------------------------

    ws.Range("A1:Q1").Font.Bold = True
    ws.Range("O2:O4").Font.Bold = True
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = ws.Range("I1").Value
    ws.Range("Q1").Value = "Value"
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"

'Variables--------------------------------------------------------

    Dim Ticker_Symbol As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim LastRow As Long
    Dim Total_Stock_Volume As Double
    Dim Summary_Table_Row As Long
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Old_Amount As Long
    Dim i As Long
    
    Total_Stock_Volume = 0
    Summary_Table_Row = 2
    Old_Amount = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Logic (start if/then)---------------------------------------------

    For i = 2 To LastRow
    
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
'Ticker Symbol------------------------------------------------------

            Ticker_Symbol = ws.Cells(i, 1).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
                
'Total Stock Volume---------------------------------------------------
                
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            Total_Stock_Volume = 0

'Yearly Change--------------------------------------------------------

            Open_Price = ws.Range("C" & Old_Amount)
            Close_Price = ws.Range("F" & i)
            Yearly_Change = Close_Price - Open_Price
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
'Percent Change-------------------------------------------------------

            If Open_Price = 0 Then
                Percent_Change = 0
            Else
                Open_Price = ws.Range("C" & Old_Amount)
                Percent_Change = Yearly_Change / Open_Price
            End If
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change

'Color Formatting---------------------------------------------------------

            If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
'Next Row--------------------------------------------------------------------

            Summary_Table_Row = Summary_Table_Row + 1
            Old_Amount = i + 1
            End If
            
        Next i

'Bonus-----------------------------------------------------------------------

    Dim LastRowBonus As Long
    LastRowBonus = Cells(Rows.Count, 11).End(xlUp).Row
            
        For i = 2 To LastRowBonus

            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
            End If

            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
            End If

            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
            End If

        Next i
            ws.Columns("A:Q").AutoFit
    Next ws

End Sub
