*/ Remove first, last x or certain position characters from text strings with ease
=RIGHT(A4, LEN(A4)-2)
=LEFT(A4, LEN(A4)-9)









*/ v_lookup
Check if one column value exists in another column using MATCH
=MATCH(lookup_value, lookup_array, [match_type])
=NOT(ISERROR(MATCH(A2,$B$2:$B$1001,0)))   
A2 : Column to be found in other column; $B$2:$B$1001 : Column in which A2 to be searched
    Lookup_value (required argument) – This is the value that we want to look up.
    Lookup_array (required argument) – The data array that is to be searched.
    Match_type (optional argument) – It can be set to 1, 0, or -1 to return results as given below:
EX : =MATCH(I10,$A$2:$A$9,0) --> 2 = Present I10->Value to be looked in array A2-A9

=IF(ISERROR(VLOOKUP(B5,$C$5:$C$33,1,FALSE)),FALSE,TRUE)

*/ Excel IF Function
 =IF(A1>70,"Pass","Fail")
=IF(D5="S","Small","Large")
 =IF(A1="10x12",120,IF(A1="8x8",64,IF(A1="6x6",36)))
Result: 120
=IF(I53="C1",5,IF(I53="C0",4,IF(I53="U1",D,IF(I53="P1",1,IF(I53="P0",0,IF(I53="B1",9,IF(I53="B0",8,1111111)))))))
 
 
 =CONCATENATE(D11,E11)
 =LEFT(D11,2)
 
 =REPLACE(F7,1,2,"KTE")
 
 C050062
 
G711	SW
C1	5
C0	4
U1	D 
U0	C 
P1	1
P0	0
B1	9
B0	8

If ( first two character is c1 then 5, elseif c0 then 4)
	
=NOT(ISERROR(MATCH(A2,$E$2:$E$801,0)))

*/ CONCATENATE






1. =IF(ISERROR(VLOOKUP(B5,$C$5:$C$33,1,FALSE)),FALSE,TRUE)
2. =NOT(ISERROR(MATCH(B5,$C$5:$C$33,0)))

=====================================*/ Word /==========================================
1. Heading Style
Menu--> Numbering(Bullets) --> Set numbering value --> Continue from previous list
