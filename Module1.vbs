Attribute VB_Name = "Module1"

Sub ExtractUniqueTickers()
    Dim workbookObj As Workbook
    Dim worksheetObj As Worksheet
    Dim rng As Range
    Dim uniqueRange As Range

    Set workbookObj = ThisWorkbook

    Dim worksheetsToProcess As Variant
    worksheetsToProcess = Array("Q1", "Q2", "Q3", "Q4")

    For i = LBound(worksheetsToProcess) To UBound(worksheetsToProcess)

        Set worksheetObj = workbookObj.Worksheets(worksheetsToProcess(i))
        Set rng = worksheetObj.Range("I1")
        Set uniqueRange = worksheetObj.Range("I1")
        
        worksheetObj.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=uniqueRange, Unique:=True

        worksheetObj.Cells(1, 9).Value = "Ticker"

    Next i
End Sub
