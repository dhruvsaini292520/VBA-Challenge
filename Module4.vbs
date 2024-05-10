Attribute VB_Name = "Module4"
Sub CalculateTickerTotalVolume()
    Dim workbookObj As Workbook
    Dim worksheetObj As Worksheet
    Dim rangeObj As Range
    Dim stockTicker As String
    Dim volumeDict As Object
    Dim totalVolume As Double
    Dim sheetsToProcess As Variant
    Dim currentRow As Long
    Dim tickerRow As Long

    Set workbookObj = ThisWorkbook

    sheetsToProcess = Array("Q1", "Q2", "Q3", "Q4")

    For i = LBound(sheetsToProcess) To UBound(sheetsToProcess)
        Set worksheetObj = workbookObj.Worksheets(sheetsToProcess(i))
        Set volumeDict = CreateObject("Scripting.Dictionary")
        
        currentRow = 2
        Do While worksheetObj.Cells(currentRow, 1).Value <> ""
            stockTicker = worksheetObj.Cells(currentRow, 1).Value
            If volumeDict.Exists(stockTicker) Then
                volumeDict(stockTicker) = volumeDict(stockTicker) + worksheetObj.Cells(currentRow, 7).Value
            Else
                volumeDict.Add stockTicker, worksheetObj.Cells(currentRow, 7).Value
            End If
            currentRow = currentRow + 1
        Loop
        
        Set rangeObj = worksheetObj.Range("I2:I" & worksheetObj.Cells(worksheetObj.Rows.Count, 9).End(xlUp).row)

        For Each cell In rangeObj
            stockTicker = cell.Value
            If volumeDict.Exists(stockTicker) Then
                cell.Offset(0, 3).Value = volumeDict(stockTicker)
            End If
        Next cell

        worksheetObj.Cells(1, 12).Value = "Total Stock Volume"
        volumeDict.RemoveAll
    Next i
    
End Sub

