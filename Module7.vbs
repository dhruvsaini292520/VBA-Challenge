Attribute VB_Name = "Module7"
Sub FindHighestTotalVolume()
    Dim workbookObj As Workbook
    Dim worksheetObj As Worksheet
    Dim sheetNames As Variant
    Dim rowNum As Long
    Dim highestValue As Double
    Dim highestTicker As String
    
    highestValue = -1E+308
    highestTicker = ""
    
    Set workbookObj = ThisWorkbook
    
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        
        Set worksheetObj = workbookObj.Worksheets(sheetNames(i))
        
        rowNum = 2
        Do While worksheetObj.Cells(rowNum, 12).Value <> ""
            Dim currentValue As Double
            currentValue = worksheetObj.Cells(rowNum, 12).Value
            
            If currentValue > highestValue Then
                highestValue = currentValue
                highestTicker = worksheetObj.Cells(rowNum, 1).Value
            End If
 
            rowNum = rowNum + 1
        Loop
    Next i
    
    workbookObj.Sheets(1).Cells(4, 16).Value = highestValue
    workbookObj.Sheets(1).Cells(4, 15).Value = highestTicker
    workbookObj.Sheets(1).Cells(4, 14).Value = "Greatest Total Volume"
    
End Sub

