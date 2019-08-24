Sub stock()

Dim Current As Worksheet

  For Each Current In Worksheets
  Current.Activate

             ' Set an initial variable for ticker
            Dim Ticker As String

            ' Set an initial variable for holding the volume
            Dim TotalVolume As Double
            TotalVolume = 0

            ' Keep track of the location for each ticker in the summary table
            Dim Summary_Table_Row As Double
            Summary_Table_Row = 2

            ' Define last row
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row

                ' Loop through all tickers
                For j = 2 To lastrow

                    ' ticker difference
                    If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then

                    ' Set the ticker name
                    Ticker = Cells(j, 1).Value

                    ' Add to the volume total
                    TotalVolume = TotalVolume + Cells(j, 7).Value

                    ' Print the ticker in the Summary Table
                    Range("I" & Summary_Table_Row).Value = Ticker

                    ' Print the volume total to the Summary Table
                    Range("J" & Summary_Table_Row).Value = TotalVolume

                    ' Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    ' Reset the Brand Total
                    TotalVolume = 0

                    ' If next cell is same
                    Else

                    ' Add to the Brand Total
                    TotalVolume = TotalVolume + Cells(j, 7).Value

                    End If

                Next j

         Next

End Sub

