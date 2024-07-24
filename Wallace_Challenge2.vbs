Sub Stock()
    Dim Stock As String
    Dim i As Long
    Dim j As Long
    Dim k As Double
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim start As Double
    Dim finish As Double
    Dim volume As Double
    Dim percentChange As Double
    Dim firstRow As Long
    Dim maxPercentChange As Double
    Dim maxStock As String
    Dim minPercentChange As Double
    Dim minStock As String
    Dim maxVolume As Double
    Dim maxVolumeStock As String

    For Each ws In Worksheets
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = "Stock Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greats % Decrease"
        ws.Cells(4, 14).Value = "Greates Total Volume"
        
        j = 2
        volume = 0
        
         maxPercentChange = -999999
         minPercentChange = 999999
         maxVolume = 0
         maxStock = ""
         minStock = ""
         maxVolumeStock = ""
        
        For i = 2 To LastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                firstRow = i
                start = ws.Cells(i, 3).Value
            End If
            
        volume = volume + ws.Cells(i, 7).Value
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                Stock = ws.Cells(i, 1).Value
                finish = ws.Cells(i, 6).Value
                
                ws.Cells(j, 9).Value = Stock
                ws.Cells(j, 10).Value = finish - start
                
            If start <> 0 Then
                    percentChange = (finish - start) / start
                Else
                    percentChange = 0
            End If
                
                ws.Cells(j, 11).Value = Round(percentChange, 4)
                ws.Cells(j, 12).Value = volume
                ws.Cells(j, 11).NumberFormat = "0.00%"

               
            If percentChange > maxPercentChange Then
                maxPercentChange = percentChange
                maxStock = Stock
                End If
                
            If percentChange < minPercentChange Then
                   minPercentChange = percentChange
                   minStock = Stock
                End If
                
                If volume > maxVolume Then
                    maxVolume = volume
                    maxVolumeStock = Stock
                End If
         
                         
              j = j + 1
               volume = 0
               
            End If
        Next i

        
For k = 2 To j - 1
            If ws.Cells(k, 11).Value >= 0 Then
                ws.Cells(k, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(k, 11).Interior.ColorIndex = 4
            End If
        Next k

                ws.Cells(2, 15).Value = maxStock
                ws.Cells(2, 15).Value = Round(maxPercentChange, 4)
                ws.Cells(2, 15).NumberFormat = "0.00%"

                ws.Cells(3, 15).Value = minStock
                ws.Cells(3, 15).Value = Round(minPercentChange, 4)
                ws.Cells(3, 15).NumberFormat = "0.00%"
                
                ws.Cells(4, 15).Value = maxVolume

    Next ws
    
End Sub
