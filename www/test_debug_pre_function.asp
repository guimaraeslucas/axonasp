<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Debug: Function pre() Call Issue</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .test { margin: 20px 0; padding: 15px; background: white; border: 1px solid #ddd; border-radius: 5px; }
        .error { background: #f8d7da; border-color: #f5c6cb; }
        .success { background: #d4edda; border-color: #28a745; }
        pre { background: #f8f9fa; padding: 10px; border-radius: 3px; overflow-x: auto; font-size: 12px; }
        code { background: #e9ecef; padding: 2px 4px; }
        h2 { color: #333; }
    </style>
</head>
<body>
    <h1>Debug: Function 'pre' - Different Calling Styles</h1>
    
    <%
    ' Define the function
    Function pre(value)
        Dim output
        output = Replace(value, vbTab, " ", 1, -1, 1)
        While InStr(output, "    ") <> 0
            output = Replace(output, "    ", "   ", 1, -1, 1)
        Wend
    End Function
    
    Dim testInput : testInput = "Hello" & vbTab & "World"
    Dim result
    
    ' Test 1: Direct assignment (should work)
    Response.Write "<div class='test success'>"
    Response.Write "<h2>Test 1: Direct Assignment</h2>"
    Response.Write "<p><code>result = pre(testInput)</code></p>"
    result = pre(testInput)
    Response.Write "<p><strong>Result:</strong> " & result & "</p>"
    Response.Write "</div>"
    
    ' Test 2: Direct Response.Write with parentheses (should work)
    Response.Write "<div class='test success'>"
    Response.Write "<h2>Test 2: Response.Write with Function Call</h2>"
    Response.Write "<p><code>Response.Write(pre(testInput))</code></p>"
    Response.Write(pre(testInput))
    Response.Write "</div>"
    
    ' Test 3: Output tag with simple argument (the problematic case)
    Response.Write "<div class='test'>"
    Response.Write "<h2>Test 3: Output Tag (<%=...%>) with Direct Argument</h2>"
    Response.Write "<p><code>&lt;%= pre(testInput) %&gt;</code></p>"
    Response.Write "<p><strong>Result:</strong> "
    %>
    <%= pre(testInput) %>
    <%
    Response.Write "</p>"
    Response.Write "</div>"
    
    ' Test 4: Output tag with complex argument (the user's actual case)
    Response.Write "<div class='test'>"
    Response.Write "<h2>Test 4: Output Tag with Method Chaining</h2>"
    Response.Write "<p><code>&lt;%= pre(UCase(testInput)) %&gt;</code></p>"
    Response.Write "<p><strong>Result:</strong> "
    %>
    <%= pre(UCase(testInput)) %>
    <%
    Response.Write "</p>"
    Response.Write "</div>"
    
    ' Test 5: Output tag with nested calls
    Response.Write "<div class='test'>"
    Response.Write "<h2>Test 5: Output Tag with Nested Method Calls</h2>"
    Response.Write "<p><code>&lt;%= pre(Server.HTMLEncode(testInput)) %&gt;</code></p>"
    Response.Write "<p><strong>Result:</strong> "
    %>
    <%= pre(Server.HTMLEncode(testInput)) %>
    <%
    Response.Write "</p>"
    Response.Write "</div>"
    %>
</body>
</html>
