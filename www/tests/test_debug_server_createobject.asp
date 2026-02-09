<% @Language=VBScript %>
<%
Response.ContentType = "text/html"
On Error Resume Next
%>
<!DOCTYPE html>
<html>
<head>
    <title>Debug Server CreateObject</title>
</head>
<body>
    <h1>Debug Server Create Object</h1>
    
    <%
    Response.Write "<h2>Test 1: CreateObject direct</h2>"
    Set obj1 = CreateObject("WScript.Shell")
    Response.Write "<p>CreateObject('WScript.Shell') = " & obj1 & "</p>"
    
    Response.Write "<h2>Test 2: Server.CreateObject</h2>"
    Set obj2 = Server.CreateObject("WScript.Shell")
    Response.Write "<p>Server.CreateObject('WScript.Shell') = " & obj2 & "</p>"
    Response.Write "<p>Err.Number = " & Err.Number & ", Err.Description = " & Err.Description & "</p>"
    
    %>
    
</body>
</html>
