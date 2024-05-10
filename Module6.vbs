Attribute VB_Name = "Module6"
Sub CalculatePercentDecrease()
    Dim workbookObj As Workbook
    Dim worksheetObj As Worksheet
    Dim sheetNames As Variant
    Dim firstPriceDict As Object
    Dim lastPriceDict As Object
    Dim stockTicker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim priceChange As Double
    Dim percentChange As Double
    Dim rowNum As Long
    Dim outputRowNum As Long
    Dim greatestPercentDecrease As Double
    Dim greatestTicker As String
  
    greatestPercentDecrease = 9999
    greatestTicker = ""
    
    Set workbookObj = ThisWorkbook
    
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        
        Set worksheetObj = workbookObj.Worksheets(sheetNames(i))
        Set firstPriceDict = CreateObject("Scripting.Dictionary")
        Set lastPriceDict = CreateObject("Scripting.Dictionary")

        rowNum = 2
        outputRowNum = 2
        
        Do While worksheetObj.Cells(rowNum, 1).Value <> ""
            stockTicker = worksheetObj.Cells(rowNum, 1).Value
            openingPrice = worksheetObj.Cells(rowNum, 3).Value
            closingPrice = worksheetObj.Cells(rowNum, 6).Value
    
            If Not firstPriceDict.Exists(stockTicker) Then
                firstPriceDict.Add stockTicker, openingPrice
            End If
    
            lastPriceDict(stockTicker) = closingPrice

            rowNum = rowNum + 1
        Loop
        
        For Each key In firstPriceDict.Keys
            priceChange = lastPriceDict(key) - firstPriceDict(key)
            percentChange = ((priceChange / firstPriceDict(key)) * 100 / 100)
            worksheetObj.Cells(outputRowNum, 11).Value = percentChange
            
            If percentChange < greatestPercentDecrease Then
                greatestPercentDecrease = percentChange
                greatestTicker = key
            End If
            
            outputRowNum = outputRowNum + 1
        Next key
        
        worksheetObj.Cells(1, 10).Value = "Quaterly Change"
        worksheetObj.Cells(1, 11).Value = "Percent Change"
    Next i

    workbookObj.Sheets(1).Cells(3, 16).Value = greatestPercentDecrease
    workbookObj.Sheets(1).Cells(3, 15).Value = greatestTicker
    workbookObj.Sheets(1).Cells(3, 14).Value = "Greatest % Decrease"
    
End Sub

