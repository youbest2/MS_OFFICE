Sub AddDashes()

Dim SrchRng As Range, cel As Range

Set SrchRng = Range("RANGE TO SEARCH")

For Each cel In SrchRng
    If InStr(1, cel.Value, "TOTAL") > 0 Then
        cel.Offset(1, 0).Value = "-"
    End If
Next cel

End Sub


'========================================================

Sub IfContains()
    If InStr(ActiveCell.Value, "string") > 0 Then
        MsgBox "The string contains the value."
    Else
        MsgBox "The string doesn't contain the value."
    End If
End Sub
