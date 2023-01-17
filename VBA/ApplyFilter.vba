Sub ApplyFilter()

' Declare variables for column index and row index
' these variables will store the column index of the "Path", "Error Number" and "Severity" columns
Dim colPath As Integer
Dim colErrorNumber As Integer
Dim colSeverity As Integer

' Search for columns with specific headers
' this loop will iterate through each column in the first row and check if the cell value matches the expected column header
For i = 1 To ActiveSheet.UsedRange.Columns.Count
    If ActiveSheet.Cells(1, i).Value = "Path" Then
        colPath = i
    ElseIf ActiveSheet.Cells(1, i).Value = "Error Number" Then
        colErrorNumber = i
    ElseIf ActiveSheet.Cells(1, i).Value = "Severity" Then
        colSeverity = i
    End If
Next i

' Apply filter for Path column
' this line will apply a filter on the specified column (colPath) with two criteria:
' the cell value contains "DiagHandler" or "DiagServices"
ActiveSheet.Range("A1").AutoFilter field:=colPath, Criteria1:="*DiagHandler*", _
    Operator:=xlOr, Criteria2:="*DiagServices*"

' Apply filter for Error Number column
' this line will apply a filter on the specified column (colErrorNumber) with one criteria:
' the cell value contains "Misra"
ActiveSheet.Range("A1").AutoFilter field:=colErrorNumber, Criteria1:="*Misra*"

End Sub
