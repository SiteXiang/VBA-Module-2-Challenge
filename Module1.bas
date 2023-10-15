Attribute VB_Name = "Module1"
Sub CalculateYearlyChange()

    Set wb = ThisWorkbook

    For Each ws In wb.Sheets
 
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
      
        If ws.Cells(1, "A").Value = "<ticker>" And _
           ws.Cells(1, "B").Value = "<date>" And _
           ws.Cells(1, "C").Value = "<open>" And _
           ws.Cells(1, "D").Value = "<high>" And _
           ws.Cells(1, "E").Value = "<low>" And _
           ws.Cells(1, "F").Value = "<close>" And _
           ws.Cells(1, "G").Value = "<vol>" Then
        
            ws.Cells(1, "I").Value = "Ticker"
            ws.Cells(1, "J").Value = "Yearly Change"
            ws.Cells(1, "K").Value = "Percent Change"
            ws.Cells(1, "L").Value = "Total Stock Volume"
       
            Set stockData = CreateObject("Scripting.Dictionary")
        
            For i = 2 To lastRow
             
                ticker = ws.Cells(i, "A").Value
                openPrice = ws.Cells(i, "C").Value
                closePrice = ws.Cells(i, "F").Value
                volume = CDbl(ws.Cells(i, "G").Value)
           
                yearlyChange = closePrice - openPrice
              
                percentChange = (yearlyChange / openPrice) * 100
             
                If stockData.Exists(ticker) Then
                 
                    existingData = stockData(ticker)
                    existingData(0) = existingData(0) + yearlyChange
                    existingData(1) = existingData(1) + percentChange
                    existingData(2) = existingData(2) + volume
                    stockData(ticker) = existingData
                Else
                   
                    Dim newData(2) As Variant
                    newData(0) = yearlyChange
                    newData(1) = percentChange
                    newData(2) = volume
                    stockData.Add ticker, newData
                End If
                
                ws.Cells(i, "I").ClearContents
                ws.Cells(i, "J").ClearContents
                ws.Cells(i, "K").ClearContents
                ws.Cells(i, "L").ClearContents
            Next i
           
            RowIndex = 2
       
            For Each tickerKey In stockData.Keys
            
                Data = stockData(tickerKey)
                ws.Cells(RowIndex, "I").Value = tickerKey
                ws.Cells(RowIndex, "J").Value = Data(0)
                ws.Cells(RowIndex, "K").Value = Data(1)
                ws.Cells(RowIndex, "K").Value = Format(Data(1), "0.00\%")
                ws.Cells(RowIndex, "L").Value = Data(2)
                RowIndex = RowIndex + 1
            Next tickerKey
        End If
    Next ws
End Sub


Sub AddSummaryToWorksheets()

    Set wb = ThisWorkbook

    For Each ws In wb.Sheets
     
        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0

        For i = 2 To lastRow
            percentChange = ws.Cells(i, "K").Value
            volume = ws.Cells(i, "L").Value
            ticker = ws.Cells(i, "I").Value

            If percentChange > maxIncrease Then
                maxIncrease = percentChange
                maxIncreaseTicker = ticker
            End If

            If percentChange < maxDecrease Then
                maxDecrease = percentChange
                maxDecreaseTicker = ticker
            End If

            If volume > maxVolume Then
                maxVolume = volume
                maxVolumeTicker = ticker
            End If
        Next i

        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
        ws.Cells(2, "P").Value = maxIncreaseTicker
        ws.Cells(2, "Q").Value = Format(maxIncrease, "0.00%")
        ws.Cells(3, "P").Value = maxDecreaseTicker
        ws.Cells(3, "Q").Value = Format(maxDecrease, "0.00%")
        ws.Cells(4, "P").Value = maxVolumeTicker
        ws.Cells(4, "Q").Value = maxVolume
    Next ws
End Sub


