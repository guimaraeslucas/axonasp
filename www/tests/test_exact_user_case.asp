<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - User Case Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        p { margin: 15px 0; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - User Case Test</h1>
        <div class="intro">
            <p>Tests specific user-reported case with function definition and string operations.</p>
        </div>
    
    <%
    Function pre(value)
        Dim pre
        pre = Replace(value, vbTab, " ", 1, -1, 1)
        While InStr(pre, "    ") <> 0
            pre = Replace(pre, "    ", "   ", 1, -1, 1)
        Wend
    End Function
    
    ' Simulating the user's call
    %>
    
    <p>Test 1: Simple call</p>
    <p><%= pre("test") %></p>
    
    <p>Test 2: With Server.HTMLEncode</p>
    <% 
    Dim testContent : testContent = "hello" & vbTab & "world"
    %>
    <p><%= pre(Server.HTMLEncode(testContent)) %></p>
    
</body>
</html>
