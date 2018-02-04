Dim MyString, MyArray, Msg

Set regEx = New RegExp            ' Create regular expression.

str1 = "ISQL>  Name                             Attribution Type"
'str1 = " K_ACCESS_CODE                    PRIMARY KEY CHAR(6)"

arr1 = split(str1, "ISQL>")

str1 = trim(arr1(1))

regEx.Pattern = "(\S)\s+"            ' Set pattern.

regEx.IgnoreCase = True            ' Make case insensitive.

Msg = regEx.Replace(str1, "$1 ")   ' Make replacement.

MyArray = Split(Msg)

Wscript.Echo "Ubound=" & UBound(MyArray) & vbCrLf & Msg

for i=0 to UBound(MyArray)
	Wscript.Echo "MyArray(" & i & ")=" & MyArray(i)
next
