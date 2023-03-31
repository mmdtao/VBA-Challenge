Attribute VB_Name = "Module1"
Sub tickerInfo()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
        Dim tickerName As String
        
    
        Dim tickerOpen As Double
        Dim tickerClose As Double
        tickerOpen = 0
        tickerClose = 0
    
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        Dim totalStockVolume As Double
        totalStockVolume = 0
    
        Cells(1, 17).Value = "Value"
        Cells(1, 16).Value = "Ticker"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total StockVolume"
        
        With ws
        
            Dim i As Integer

            
            For i = 2 To 22771
        
                If Cells(i, 1).Value = Cells(i + 1, 1).Value And Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                    tickerOpen = Cells(i, 3).Value
                End If
            
                If Cells(i, 6).Value = Cells(i - 1, 6).Value And Cells(i + 1, 6).Value <> Cells(i, 6).Value Then
                    tickerClose = Cells(i, 3).Value
                End If
                
        
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            
                    tickerName = Cells(i, 1).Value
                    totalStockVolume = totalStockVolume + Cells(i, 7).Value
                
                
                    Range("I" & Summary_Table_Row).Value = tickerName
                    Range("L" & Summary_Table_Row).Value = totalStockVolume
                    Range("J" & Summary_Table_Row).Value = (tickerOpen - tickerClose)
                    Range("K" & Summary_Table_Row).Value = FormatPercent((tickerOpen - tickerClose) / tickerOpen)
                
                
                    Summary_Table_Row = Summary_Table_Row + 1
                
                    totalStockVolume = 0
                    tickerOpen = 0
                    tickerClose = 0
                
                Else
                    totalStockVolume = totalStockVolume + Cells(i, 7).Value
                End If
                
                If Cells(i, 11).Value > 0 Then
                        Cells(i, 11).Interior.ColorIndex = 4
                Else
                        Cells(i, 11).Interior.ColorIndex = 3
                End If
                
        End With
    Next ws


End Sub
