<%
Response.Write "<h1>Testing ExecuteGlobal</h1>"

' Create a simple function
Dim code
Dim q
q = Chr(34)
code = "Function TestFunc(x)" & vbCrLf & _
       "TestFunc = x & " & q & " processed" & q & vbCrLf & _
       "End Function"

Response.Write "<p>Executing code:</p><pre>" & Server.HTMLEncode(code) & "</pre>"

ExecuteGlobal code

Response.Write "<p>Calling TestFunc('hello'): " & TestFunc("hello") & "</p>"

Response.Write "<p>Test complete</p>"
%>
