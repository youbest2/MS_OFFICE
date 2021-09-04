Sub FinalMacroColor()
    
    Dim i_Row       As Integer
    Dim J_Column    As Integer
    Dim var         As Double
    
    Dim NextRow     As Integer
    Dim MaxColumn   As Integer
    
    MaxColumn = 26
    'Cells(Row_num, Col_num)
    For i_Row = 2 To 11
        For J_Column = 4 To 26
            var = Cells(i_Row, 3).Value
            
            If Cells(i_Row, J_Column).Value > var And J_Column <= MaxColumn Then
                Cells(i_Row, J_Column).Interior.Color = VBA.ColorConstants.vbGreen
            Else
                Cells(i_Row, J_Column).Interior.Color = VBA.ColorConstants.vbRed
            End If
            If J_Column = MaxColumn Then
                Exit For
            End If
            
        Next J_Column
    Next i_Row
    
End Sub
