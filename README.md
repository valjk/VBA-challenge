Xpert Learning Assistant(XLA) - My initial code and output worked for each Ticker row, and I needed assistance with outputting the summary table.
XLA recommended adding the "Ticker Check" code above the code I wrote for "Update totals for summary" and "Output results for each ticker", then it worked.

  ' Ticker Check
            Dim found As Boolean
            found = False

            For j = 0 To tickerCount - 1
                If tickers(j) = ticker Then
                    quarterlyChanges(j) = closingPrice - firstOpeningPrices(j)
                    If firstOpeningPrices(j) <> 0 Then
                        percentageChanges(j) = (quarterlyChanges(j) / firstOpeningPrices(j)) * 100
                    Else
                        percentageChanges(j) = 0
                    End If
                    totalVolumes(j) = totalVolumes(j) + totalVolume
                    found = True
                    Exit For
                End If
            Next j

            If Not found Then
                ReDim Preserve tickers(tickerCount)
                ReDim Preserve quarterlyChanges(tickerCount)
                ReDim Preserve percentageChanges(tickerCount)
                ReDim Preserve totalVolumes(tickerCount)
                ReDim Preserve firstOpeningPrices(tickerCount)

                tickers(tickerCount) = ticker
                firstOpeningPrices(tickerCount) = openingPrice
                quarterlyChanges(tickerCount) = 0
                percentageChanges(tickerCount) = 0
                totalVolumes(tickerCount) = totalVolume
                tickerCount = tickerCount + 1
            End If
    
