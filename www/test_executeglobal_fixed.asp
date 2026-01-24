<%
Response.Write "<h1>Testing ExecuteGlobal - Corrected</h1>"

' Create a simple function with proper quotes
Dim code
code = "Function TestFunc(x)" & vbCrLf & _
       "TestFunc = x & "" processed""" & vbCrLf & _
       "End Function"

Response.Write "<p>Executing code:</p><pre>" & Server.HTMLEncode(code) & "</pre>"

ExecuteGlobal code

Response.Write "<p>Calling TestFunc('hello'): " & TestFunc("hello") & "</p>"

Response.Write "<p>Test complete</p>"
%>

