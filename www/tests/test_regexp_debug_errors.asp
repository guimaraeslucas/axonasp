<%@ Language="VBScript" %>
<html>
<head>
    <title>RegExp Test - Error Debugging</title>
    <style>
        body { font-family: Arial; margin: 20px; }
        .error { color: red; background: #ffe0e0; padding: 10px; margin: 10px 0; border: 1px solid red; }
        .info { color: blue; background: #e0f0ff; padding: 10px; margin: 10px 0; border: 1px solid blue; }
        code { background: #f0f0f0; padding: 5px; font-family: monospace; }
    </style>
</head>
<body>

<h1>RegExp Test 1 - Debug Issues</h1>

<%
    Response.Write "<h2>Attempting Test 1 (Object Creation)</h2>"
    
    On Error Resume Next
    
    ' Try Test 1
    Response.Write "<div class='info'>"
    Response.Write "<strong>Step 1:</strong> Server.CreateObject('G3REGEXP')<br>"
    
    Set regex = Server.CreateObject("G3REGEXP")
    
    if Err.Number <> 0 then
        Response.Write "<div class='error'>ERROR: " & Err.Number & " - " & Err.Description & "</div>"
    else
        Response.Write "Success! Object type: " & TypeName(regex) & "<br>"
        Response.Write "Object is: " & CStr(regex) & "<br>"
    end if
    Response.Write "</div>"
    
    ' Try to access pattern
    Response.Write "<div class='info'>"
    Response.Write "<strong>Step 2:</strong> Try setting Pattern via direct assignment<br>"
    
    Err.Clear
    regex.Pattern = "\d+"
    
    if Err.Number <> 0 then
        Response.Write "<div class='error'>ERROR: " & Err.Number & " - " & Err.Description & "</div>"
    else
        Response.Write "Success! Pattern set<br>"
    end if
    Response.Write "</div>"
    
    ' Try to call Test()
    Response.Write "<div class='info'>"
    Response.Write "<strong>Step 3:</strong> Call Test('123')<br>"
    
    Err.Clear
    result = regex.Test("123")
    
    if Err.Number <> 0 then
        Response.Write "<div class='error'>ERROR: " & Err.Number & " - " & Err.Description & "</div>"
    else
        Response.Write "Success! Result: " & result & "<br>"
    end if
    Response.Write "</div>"
    
    ' Try Execute()
    Response.Write "<div class='info'>"
    Response.Write "<strong>Step 4:</strong> Call Execute('abc123def456')<br>"
    
    Err.Clear
    Set matches = regex.Execute("abc123def456")
    
    if Err.Number <> 0 then
        Response.Write "<div class='error'>ERROR: " & Err.Number & " - " & Err.Description & "</div>"
    else
        Response.Write "Success! Matches object: " & TypeName(matches) & "<br>"
        
        if Not (matches Is Nothing) then
            Err.Clear
            count = matches.Count
            if Err.Number <> 0 then
                Response.Write "<div class='error'>ERROR accessing .Count: " & Err.Number & " - " & Err.Description & "</div>"
            else
                Response.Write "Matches.Count: " & count & "<br>"
            end if
        end if
    end if
    Response.Write "</div>"
%>

</body>
</html>
