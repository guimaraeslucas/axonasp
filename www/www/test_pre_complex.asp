<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Test: pre() with aspl.loadText simulation</title>
</head>
<body>
    <h1>Test Pre with Complex Arguments</h1>
    
    <%
    ' Simple loadText simulation
    Function loadText(path)
        ' Simulate reading file content
        loadText = "Hello World With Tabs" & vbTab & "More Content"
    End Function
    
    Function pre(value)
        Dim output
        output = Replace(value, vbTab, " ", 1, -1, 1)
        While InStr(output, "    ") <> 0
            output = Replace(output, "    ", "   ", 1, -1, 1)
        Wend
    End Function
    
    %>
    
    <p>Test 1: Direct pre() call</p>
    <p><%= pre("test" & vbTab & "value") %></p>
    
    <p>Test 2: pre() with Server.HTMLEncode</p>
    <p><%= pre(Server.HTMLEncode("test" & vbTab & "value")) %></p>
    
    <p>Test 3: pre() with loadText() function</p>
    <p><%= pre(loadText("test.txt")) %></p>
    
    <p>Test 4: pre() with Server.HTMLEncode(loadText())</p>
    <p><%= pre(Server.HTMLEncode(loadText("test.txt"))) %></p>
    
</body>
</html>
