=IF(AND(B2>50, B2>50), "Pass", "Fail")
=IF(AND(B2="pass", C2="pass"), "Good!", "Bad")
Important note! The AND function checks all the conditions, even if the already tested one(s) evaluated to FALSE. Such behavior is a bit unusual since in most of programming languages, subsequent conditions are not tested if any of the previous tests has returned FALSE.

In practice, a seemingly correct IF statement may result in an error because of this specificity. For example, the below formula would return #DIV/0! ("divide by zero" error) if cell A2 is equal to 0:

=IF(AND(A2<>0, (1/A2)>0.5),"Good", "Bad")

The avoid this, you should use a nested IF function:

=IF(A2<>0, IF((1/A2)>0.5, "Good", "Bad"), "Bad")

=IF(OR(B2>50, B2>50), "Pass", "Fail")

IF with multiple AND & OR statements
If your task requires evaluating several sets of multiple conditions, you will have to utilize both AND & OR functions at a time.

In our sample table, suppose you have the following criteria for checking the exam results:

Condition 1: exam1>50 and exam2>50
Condition 2: exam1>40 and exam2>60
If either of the conditions is met, the final exam is deemed passed.

At first sight, the formula seems a little tricky, but in fact it is not! You just express each of the above conditions as an AND statement and nest them in the OR function (since it's not necessary to meet both conditions, either will suffice):

OR(AND(B2>50, C2>50), AND(B2>40, C2>60)

Nested IF statement to check multiple logical tests
=IF(B2>=60, "Good", IF(B2>40, "Satisfactory", "Poor"))


Excel IF array formula with multiple conditions
To evaluate conditions with the AND logic, use the asterisk:
IF(condition1) * (condition2) * …, value_if_true, value_if_false)
To test conditions with the OR logic, use the plus sign:
IF(condition1) + (condition2) + …, value_if_true, value_if_false)                         

For example, to get "Pass" if both B2 and C2 are greater than 50, the formula is:
=IF((B2>50) * (C2>50), "Pass", "Fail")

https://www.ablebits.com/office-addins-blog/excel-if-function-multiple-conditions/

