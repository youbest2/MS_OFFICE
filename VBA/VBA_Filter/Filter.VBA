Sub ApplyFilter()
'
' Macro1 Macro
'
    ' Select cell D1
    Range("D1").Select
    ' Apply filter on the selected cell
    Selection.AutoFilter

    ' Apply filter on the range $A$1:$J$297871 for the 4th column (column D)
    ' with criteria that the cell value contains either "DiagHandler" or "DiagServices"
    ActiveSheet.Range("$A$1:$J$297871").AutoFilter Field:=4, Criteria1:= _
        "=*DiagHandler*", Operator:=xlOr, Criteria2:="=*DiagServices*"
    ' Scroll to column 3
    ActiveWindow.ScrollColumn = 3

    ' Select cell G1
    Range("G1").Select
    ' Apply filter on the range $A$1:$J$297871 for the 7th column (column G)
    ' with criteria that the cell value is one of the specified values (high, low, mandatory, medium, required)
    ActiveSheet.Range("$A$1:$J$297871").AutoFilter Field:=7, Criteria1:=Array( _
        "high", "low", "mandatory", "medium", "required"), Operator:=xlFilterValues
    ' Scroll to column 2
    ActiveWindow.ScrollColumn = 2

    ' Apply filter on the range $A$1:$J$297871 for the 1st column (column A)
    ' with criteria that the cell value contains "Misra"
    ActiveSheet.Range("$A$1:$J$297871").AutoFilter Field:=1, Criteria1:= _
         "=*Misra*"
    'End of the Sub
End Sub
