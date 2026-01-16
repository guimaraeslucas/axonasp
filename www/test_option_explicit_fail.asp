<%
Option Explicit
Response.Write "<h3>Testing Option Explicit (Expect Failure)</h3>"
Response.Write "This script should fail with 'Variable is undefined: x'<br>"
x = 10 ' Undeclared variable
Response.Write "x is " & x
%>