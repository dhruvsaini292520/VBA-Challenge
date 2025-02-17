Attribute VB_Name = "Module2"

Sub CalculateQuarterlyChange()
    Dim workbookObj As Workbook
    Dim worksheetObj As Worksheet
    Dim sheetNames As Variant
    Dim firstPriceDict As Object
    Dim lastPriceDict As Object
    Dim stockTicker As String
    Dim initialPrice As Double
    Dim closingPrice As Double
    Dim priceChange As Double
    Dim dataRowNum As Long
    Dim outputRowNum As Long

    Set workbookObj = ThisWorkbook
    
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set worksheetObj = workbookObj.Worksheets(sheetNames(i))
        Set firstPriceDict = CreateObject("Scripting.Dictionary")
        Set lastPriceDict = CreateObject("Scripting.Dictionary")

        dataRowNum = 2
        outputRowNum = 2
        
        Do While worksheetObj.Cells(dataRowNum, 1).Value <> ""
            stockTicker = worksheetObj.Cells(dataRowNum, 1).Value
            initialPrice = worksheetObj.Cells(dataRowNum, 3).Value
            closingPrice = worksheetObj.Cells(dataRowNum, 6).Value

            If Not firstPriceDict.Exists(stockTicker) Then
                firstPriceDict.Add stockTicker, initialPrice
            End If

            lastPriceDict(stockTicker) = closingPrice
            dataRowNum = dataRowNum + 1
        Loop

        For Each stockKey In firstPriceDict.Keys
            priceChange = lastPriceDict(stockKey) - firstPriceDict(stockKey)
            worksheetObj.Cells(outputRowNum, 10).Value = priceChange
            outputRowNum = outputRowNum + 1
        Next stockKey
        worksheetObj.Cells(1, 10).Value = "Quarterly Change"
    Next i
    
End Sub
