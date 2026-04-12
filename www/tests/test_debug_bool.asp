<% @Language=VBScript %>
<%
Response.ContentType = "text/html"
On Error Resume Next
%>
<!DOCTYPE html>
<html>
<head>
    <title>Debug Bool Values</title>
</head>
<body>
    <h1>Debug Bool Values</h1>
    
    <%
    ' Test basic boolean operations
    Response.Write "<h2>Test 1: Boolean literals</h2>"
    b1 = True
    b2 = False
    Response.Write "<p>True = " & b1 & "</p>"
    Response.Write "<p>False = " & b2 & "</p>"
    Response.Write "<p>Not True = " & (Not True) & "</p>"
    Response.Write "<p>Not False = " & (Not False) & "</p>"
    
    Response.Write "<h2>Test 2: Is operator results</h2>"
    Set obj = CreateObject("WScript.Shell")
    Result1 = (obj Is Nothing)
    Response.Write "<p>obj Is Nothing = " & Result1 & "</p>"
    
    Response.Write "<h2>Test 3: Not on Is result</h2>"
    Result2 = Not (obj Is Nothing)
    Response.Write "<p>Not (obj Is Nothing) = " & Result2 & "</p>"
    Response.Write "<p>Type check: Is it 0 or -1? " & Result2 & "</p>"
    
    %>
    
</body>
</html>
