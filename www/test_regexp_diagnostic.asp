<%@ Language="VBScript" %>
<html>
<head><title>RegExp - Diagnostic Test</title>
<style>
body { font-family: monospace; margin: 20px; }
.test { padding: 10px; margin: 5px 0; border: 1px solid #ccc; }
pre { background: #f0f0f0; padding: 10px; overflow-x: auto; }
</style>
</head>
<body>

<h1>RegExp Diagnostic Test</h1>

<%
    On Error Resume Next
    
    Response.Write "<h2>Test 1: Basic Object Creation</h2>"
    Response.Write "<div class='test'>"
    
    Set regex = New RegExp
    
    Response.Write "Err.Number = " & Err.Number & "<br>"
    Response.Write "Err.Description = " & Err.Description & "<br>"
    Response.Write "TypeName(regex) = " & TypeName(regex) & "<br>"
    Response.Write "IsObject(regex) = " & IsObject(regex) & "<br>"
    
    if regex Is Nothing then
        Response.Write "regex Is Nothing = TRUE<br>"
    else
        Response.Write "regex Is Nothing = FALSE<br>"
    end if
    
    Response.Write "</div>"
    
    Response.Write "<h2>Test 2: Accessing Pattern Property</h2>"
    Response.Write "<div class='test'>"
    
    Err.Clear
    Response.Write "Before: Err.Number = " & Err.Number & "<br>"
    
    regex.Pattern = "\d+"
    
    Response.Write "After assignment: Err.Number = " & Err.Number & "<br>"
    Response.Write "Err.Description = " & Err.Description & "<br>"
    
    Err.Clear
    pat = regex.Pattern
    
    Response.Write "After read: Err.Number = " & Err.Number & "<br>"
    Response.Write "Pattern value = " & pat & "<br>"
    
    Response.Write "</div>"
    
    Response.Write "<h2>Test 3: Execute() Method</h2>"
    Response.Write "<div class='test'>"
    
    Err.Clear
    Set matches = regex.Execute("abc123def")
    
    Response.Write "Err.Number = " & Err.Number & "<br>"
    Response.Write "Err.Description = " & Err.Description & "<br>"
    Response.Write "TypeName(matches) = " & TypeName(matches) & "<br>"
    Response.Write "matches Is Nothing = " & (matches Is Nothing) & "<br>"
    
    if Not (matches Is Nothing) then
        Err.Clear
        Response.Write "Trying to access matches.Count...<br>"
        cnt = matches.Count
        Response.Write "Err.Number = " & Err.Number & "<br>"
        Response.Write "Err.Description = " & Err.Description & "<br>"
        Response.Write "Count value = " & cnt & "<br>"
    end if
    
    Response.Write "</div>"
    
    Response.Write "<h2>Test 4: Direct Method Call Test()</h2>"
    Response.Write "<div class='test'>"
    
    Err.Clear
    result = regex.Test("123")
    
    Response.Write "Err.Number = " & Err.Number & "<br>"
    Response.Write "Err.Description = " & Err.Description & "<br>"
    Response.Write "Test result = " & result & "<br>"
    
    Response.Write "</div>"
%>

</body>
</html>
