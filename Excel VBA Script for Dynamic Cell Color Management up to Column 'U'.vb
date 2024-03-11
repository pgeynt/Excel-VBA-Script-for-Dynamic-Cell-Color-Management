Private prevRange As Range

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Dim Cell As Range
    Dim RowHasLavender As Boolean
    Dim LastCol As Long
    RowHasLavender = False
    
  
    LastCol = 21 
    
    For Each Cell In Target.EntireRow.Cells
        If Cell.Column > LastCol Then Exit For ' 
        
        If Cell.Interior.Color = RGB(230, 230, 250) Then ' Lavender
            RowHasLavender = True
            Exit For
        End If
    Next Cell

    If Not prevRange Is Nothing And Not RowHasLavender Then
        For Each Cell In prevRange.Cells
            If Cell.Column > LastCol Then Exit For ' "
            
            If Cell.Interior.Color = RGB(230, 230, 250) Then ' Lavender
                Cell.Interior.ColorIndex = -4142 '
            End If
        Next Cell
    End If

    If RowHasLavender Then
        Set prevRange = Target.EntireRow
        Exit Sub
    End If

    For Each Cell In Target.EntireRow.Cells
        If Cell.Column > LastCol Then Exit For ' 
        
        If Cell.Interior.ColorIndex = xlNone Or Cell.Interior.Color = RGB(255, 255, 255) Then ' vbWhite
            Cell.Interior.Color = RGB(230, 230, 250) ' Lavender
        End If
    Next Cell

    Set prevRange = Target.EntireRow
End Sub