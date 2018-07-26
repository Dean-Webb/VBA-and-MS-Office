Some great work by Raghy R
[How to extract initials from names in Excel?](http://quadexcel.com/how-to-extract-initials-from-names-in-excel/)





**I have copied this from his site:

When you are working with loads of consumer records you might need to get the initials from a given name, like TCS for Tata Consultancy Service and JFK for John F Kennedy.

We can get the initials using combining couple of text formulas LEFT(), MID(), FIND() with ISERROR().

For example cell A2 has the full name, use below formula to get the initials from full name.

**Solution #1: Use Formula
```
=LEFT(B3,1)&IF(ISERROR(FIND(" ",B3,1)),"",MID(B3,FIND(" ",B3,1)+1,1))
&IF(ISERROR(FIND(" ",B3,FIND(" ",B3,1)+1)),"",MID(B3,FIND(" ",B3,FIND(" ",B3,1)+1)+1,1))
```
To make it simple formula is limited for the names with max three parts, in case of names having more than 3 parts it will return only first 3 initials. If your name got more than 3 parts, check for solution #2

Extract Initials From Names01

** Solution #2: Use Macro/VBA Code

How about making the above formula short like =Initials(FullName) returns the initials

Extract Initials From Names02

Copy below code to your excel workbook.

Press Alt + F11 to open the Microsoft Visual Basic for Applications window.
In the pop-up window, clickInsert> Module, then paste the following VBA code into the module.
VBA: Extract initials from names
```
Function Initials(str As String) As String
    Dim sTemp() As String
    Dim i As Long
    sTemp = Split(str)
    
    For i = 0 To UBound(sTemp)
        If UCase(sTemp(i)) Like "[A-Z]*" Then
            Initials = Initials & UCase(Left(sTemp(i), 1))
        End If
    Next i

End Function
```
