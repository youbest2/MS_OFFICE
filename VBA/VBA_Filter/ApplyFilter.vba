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





Sub ApplyFilterNew()

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

' Apply filter for Severity Column
' this line will apply a filter on the specified column (colSeverity) with five criteria:
' with criteria that the cell value is one of the specified values (high, low, mandatory, medium, required)
ActiveSheet.Range("A1").AutoFilter field:=colSeverity, Criteria1:=Array( _
        "high", "low", "mandatory", "medium", "required"), Operator:=xlFilterValues
End Sub


'Criteria1:=Array(_ and Criteria1:="*DiagHandler*" are two different ways to specify filter criteria in VBA.

'Criteria1:=Array(_ is used when the filter criteria include multiple values. The Array function is used to create an array of values, and the filter will include rows where the specified column matches any of the values in the array. For example, in the code you provided, the filter criteria is Criteria1:=Array("high", "low", "mandatory", "medium", "required"), which means that the filter will include rows where the value in the 7th column is either "high", "low", "mandatory", "medium" or "required".

'Criteria1:="*DiagHandler*" is used when the filter criteria is based on a pattern or text. The Criteria1 parameter is used to specify the filter criteria and the filter will include rows where the specified column matches the criteria. The Operator:=xlOr is used to combine the two criteria, so that the filter will include rows where the value in the 4th column contains either "DiagHandler" or "DiagServices".

'In summary, Criteria1:=Array(_ is used to filter based on multiple values and Criteria1:="*DiagHandler*" is used to filter based on a pattern or text.

