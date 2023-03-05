Sub vba_add_row()
    Application.DisplayAlerts = False
    Dim shtPricing As Worksheet
    Set shtPricing = Sheets("Sheet1")
    
    Dim startingCell As Integer
    Dim endingCell As Integer
    
    startingCell = 2
    Dim previousCell As Integer
    previousCell = 4
    
    For rowCount = 2 To 1432 Step 1
        Dim cellValue As Integer
        Dim r1 As Range
        
        Dim compareResult As Integer
        
        With shtPricing
            Set r1 = Range(.Cells(rowCount, 1), .Cells(rowCount, 1))
        End With
        cellValue = r1.Value
        
        'cells are the same
        If previousCell = cellValue Then
            endingCell = rowCount
        'cells are not the same
        Else
            'merge the needed cells

            With shtPricing
                Set r1 = Range(.Cells(startingCell, 1), .Cells(endingCell, 1))
            End With
            r1.Merge False
            'merge the needed cells

            'do the border for the last line
            With shtPricing
                Set r1 = Range(.Cells(endingCell, 1), .Cells(endingCell, 21))
            End With
            r1.Borders(xlEdgeBottom) _
                .LineStyle = XlLineStyle.xlContinuous
            'do the border for the last line
            previousCell = cellValue
            startingCell = rowCount
            endingCell = rowCount
        End If
    Next rowCount
    
End Sub