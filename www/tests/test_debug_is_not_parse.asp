<% @Language=VBScript %>
<%
Response.ContentType = "text/html"
On Error Resume Next
%>
<!DOCTYPE html>
<html>
<head>
    <title>Debug Is Not Parsing</title>
</head>
<body>
    <h1>Debug Is Not Parsing</h1>
    
    <%
    Set objShell = CreateObject("WScript.Shell")
    
    Response.Write "<h2>Test: Expression evaluation</h2>"
    Response.Write "<p>objShell = " & objShell & "</p>"
    
    ' Test different ways of checking
    Result1 = (objShell Is Nothing)
    Result2 = (objShell Is Not Nothing)
    Result3 = Not (objShell Is Nothing)
    
    Response.Write "<p>objShell Is Nothing = " & Result1 & "</p>"
    Response.Write "<p>objShell Is Not Nothing = " & Result2 & "</p>"
    Response.Write "<p>Not (objShell Is Nothing) = " & Result3 & "</p>"
    
    ' Test with undefined variable
    Response.Write "<h2>Test: Undefined variable</h2>"
    Result4 = (undefinedVar Is Nothing)
    Result5 = (undefinedVar Is Not Nothing)
    Result6 = Not (undefinedVar Is Nothing)
    
    Response.Write "<p>undefinedVar Is Nothing = " & Result4 & "</p>"
    Response.Write "<p>undefinedVar Is Not Nothing = " & Result5 & "</p>"
    Response.Write "<p>Not (undefinedVar Is Nothing) = " & Result6 & "</p>"
    
    %>
    
</body>
</html>
