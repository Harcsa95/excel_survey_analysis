Sub vba_add_row()
    Application.DisplayAlerts = False
    Dim rowCount As Integer
    Dim shtPricing As Worksheet
    Set shtPricing = Sheets("Sheet1")
    
    rowCount = 2
    
    For Times = 0 To 338 Step 1
        Dim cellValue As String
        Dim cellValueTitleList() As String
        Dim arrayLength As Integer
        Dim arrayLength1 As Integer
        Dim r1 As Range
        
        With shtPricing
            Set r1 = Range(.Cells(rowCount, 1), .Cells(rowCount, 1))
        End With
        cellValue = r1.Value
        cellValueTitleList = Split(cellValue, ",")
        
        arrayLength = WorksheetFunction.CountA(cellValueTitleList)
        
        'copy rows
        For i = rowCount To rowCount + arrayLength - 2 Step 1
            With shtPricing
                Set r1 = Range(.Cells(rowCount, 1), .Cells(rowCount, 21))
            End With
    
            r1.Copy
            r1.Insert xlShiftDown
        Next i
        'copy rows
        
            
        'set the title to a single entry
        For i = 0 To arrayLength - 1 Step 1
            Cells(rowCount + i, 1).Value = LTrim(cellValueTitleList(i))
        Next i
        'set the title to a single entry

    rowCount = rowCount + arrayLength
    Next Times
End Sub
