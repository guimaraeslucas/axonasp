<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head><title>Response End Test</title></head>
<body>
    <h1>Response End Test</h1>
    <p>Line before End</p>
    <% 
    Dim d
    d = #1/1/2020#
    Response.Write "<p>Date Literal: " & d & "</p>"
    Response.End 
    %>
    <p>Line after End (Should NOT appear)</p>
</body>
</html>
