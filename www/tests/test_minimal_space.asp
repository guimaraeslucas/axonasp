<%
@ Language = VBScript
%>
<html>
    <body>
        <h1>Minimal Test</h1>
        <pre>
<%
Response.Write "Test 1: Simple output" & vbCrLf
Response.Write "Hello World" & vbCrLf

Response.Write vbCrLf & "Test 2: Space function" & vbCrLf
Dim result
result = Space(5)
Response.Write "Space(5) returned type: " & TypeName(result) & vbCrLf
Response.Write "Space(5) value: [" & result & "]" & vbCrLf

Response.Write vbCrLf & "Test 3: Done" & vbCrLf
%>
</pre>
    </body>
</html>
