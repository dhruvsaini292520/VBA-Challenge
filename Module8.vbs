Attribute VB_Name = "Module8"
Sub ColorPercentChange()
    Dim workbookObj As Workbook
    Dim worksheetObj As Worksheet
    Dim cell As Range

    Set workbookObj = ThisWorkbook
    For Each worksheetObj In workbookObj.Worksheets
       
        For Each cell In worksheetObj.Range("J2", worksheetObj.Cells(worksheetObj.Rows.Count, 10).End(xlUp))
            
            If cell.Value > 0 Then
                cell.Interior.Color = RGB(0, 255, 0)
            ElseIf cell.Value < 0 Then
                cell.Interior.Color = RGB(255, 0, 0)
            Else
                cell.Interior.ColorIndex = xlNone
            End If
        Next cell
    Next worksheetObj
End Sub

