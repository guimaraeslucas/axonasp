<% 
' Test file for Debug ASP parser functionality
' This file contains various ASP code to demonstrate debug output

Response.Write("Testing Debug ASP Parser<br>")

' Basic variable assignment
Dim x, y, z
x = 10
y = 20
z = x + y

Response.Write("Result: " & z & "<br>")

' Function declaration
Function Add(a, b)
	Add = a + b
End Function

Response.Write("Function result: " & Add(5, 3) & "<br>")

' Conditional statement
If x > 5 Then
	Response.Write("X is greater than 5<br>")
End If

' Loop
For i = 1 To 3
	Response.Write("Loop iteration: " & i & "<br>")
Next

Response.Write("Debug test completed successfully")
%>
