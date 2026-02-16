<% @Language=VBScript %>
<%
Response.ContentType = "text/html"
On Error Resume Next
%>
<!DOCTYPE html>
<html>
<head>
    <title>Debug Nothing</title>
</head>
<body>
    <h1>Debug Nothing Keyword</h1>
    
    <%
    Set objShell = CreateObject("WScript.Shell")
    
    Response.Write "<h2>Test 1: Check Against Nothing</h2>"
    Response.Write "<p>objShell = " & objShell & "</p>"
    Response.Write "<p>objShell Is Nothing = " & (objShell Is Nothing) & "</p>"
    Response.Write "<p>objShell Is Not Nothing = " & (objShell Is Not Nothing) & "</p>"
    
    Response.Write "<h2>Test 2: Undefined Variable Against Nothing</h2>"
    Response.Write "<p>undefinedVar Is Nothing = " & (undefinedVar Is Nothing) & "</p>"
    Response.Write "<p>undefinedVar Is Not Nothing = " & (undefinedVar Is Not Nothing) & "</p>"
    
    Response.Write "<h2>Test 3: Null value</h2>"
    nullVal = Null
    Response.Write "<p>nullVal = " & nullVal & "</p>"
    Response.Write "<p>nullVal Is Nothing = " & (nullVal Is Nothing) & "</p>"
    Response.Write "<p>nullVal Is Not Nothing = " & (nullVal Is Not Nothing) & "</p>"
    
    %>
    
</body>
</html>
