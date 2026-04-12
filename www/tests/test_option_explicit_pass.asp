<%
Option Explicit
Response.Write "<h3>Testing Option Explicit (Expect Success)</h3>"

Dim x
x = 10
Response.Write "x is " & x & "<br>"

Dim arr
ReDim arr(5)
arr(0) = "Hello"
Response.Write "arr(0) is " & arr(0) & "<br>"

Dim i
For i = 1 To 5
    Response.Write i & " "
Next
Response.Write "<br>Loop Done"
%>