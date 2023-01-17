This line creates an array of values and assigns it to the Criteria1 parameter of the AutoFilter method. The filter will include rows where the value in the specified column is one of the values in the array.

It is important to note that when you use Criteria1:=Array(_ you should use Operator:=xlFilterValues to apply the filter.

The =_ is not necessary when you are using Criteria1:="*DiagHandler*" because it is not an array of values and you don't need to continue the line in the next one.
