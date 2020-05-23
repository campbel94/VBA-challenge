Sub MultiYearStockSummary()
    
    ' Set variable for holding ticker symbol
    Dim TickerSymbol As String
    
    'Set variable for holding total volume of each stock
    Dim Volume As Double
    Volume = 0
    
    ' Track the location for each stock in the table
    Dim TableRow As Integer
    
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
    
    ' Track the row that the stocks opening price is in
    Dim Opening As Double
    
    'Track the row that the stocks closing price is in
    Dim Closing As Double
    
    Dim RowCounter As Integer
    RowCounter = 0
    
    ' Determine the last row of data
    Dim LastRow As Double
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
    ' Loop through each row (date) of stock data
    For i = 2 To LastRow
    
        ' Check if we are still in the same stock...
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            
            ' Add the volume
            Volume = Volume + Cells(i, 7)
            
            ' Add to hte RowCounter
            RowCounter = RowCounter + 1
            
        Else
            
            ' Set the ticker symbol
            TickerSymbol = Cells(i, 1).Value
            
            ' Add to the total volume
            Volume = Volume + Cells(i, 7).Value
            
            ' Set the opening price
            Opening = Cells(i - RowCounter, 3).Value
            
            ' Set the closing price
            Closing = Cells(i, 6).Value
            
            ' Print the tiker symbol in our summary table
            Range("I" & SummaryTableRow).Value = TickerSymbol
            
            ' Print the yearly change amount
            Range("J" & SummaryTableRow).Value = Closing - Opening
            

            ' Print the yearly change percent
            ' Cannot avoid getting "Run-time Error '6': Overflow" error
            Range("K" & SummaryTableRow).Value = Format(((Closing - Opening) / Opening), "Percent")
            
            ' Print the volume to our summary table
            Range("L" & SummaryTableRow).Value = Volume
            
                ' Apply conditional formatting
                If Range("K" & SummaryTableRow).Value < 0 Then
                    Range("K" & SummaryTableRow).Interior.ColorIndex = 3
                    
                    ElseIf Range("K" & SummaryTableRow).Value = 0 Then
                        Range("K" & SummaryTableRow).Interior.ColorIndex = 6
                    
                    ElseIf Range("K" & SummaryTableRow).Value > 0 Then
                        Range("K" & SummaryTableRow).Interior.ColorIndex = 4
                    
                End If
                
            ' Add 1 to the table row
            SummaryTableRow = SummaryTableRow + 1
            
            ' Reset the volume
            Voume = 0
            
            'Reset the Row Counter
            RowCounter = 0
        
            
        End If
        
    Next i
    
    
End Sub