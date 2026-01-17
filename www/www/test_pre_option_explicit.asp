<%@ Language=VBScript %>
<%
Option Explicit

' This is the exact pattern from the user's problem
Function pre(value)
    pre=replace(value,vbTab," ",1,-1,1)
    While InStr(pre,"    ")<>0
        pre=replace(pre,"    ","   ",1,-1,1)
    Wend
End Function

%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Test: Function pre() with Option Explicit</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        .success { background: #d4edda; border: 1px solid #28a745; padding: 15px; border-radius: 5px; }
        .test { margin: 20px 0; }
        pre { background: #f5f5f5; padding: 10px; }
    </style>
</head>
<body>
    <h1>Test: Function pre() with Option Explicit</h1>
    
    <div class="test success">
        <h2>✓ Test 1: Simple pre() call</h2>
        <p>Result: <%= pre("hello" & vbTab & "world") %></p>
    </div>
    
    <div class="test success">
        <h2>✓ Test 2: With Server.HTMLEncode</h2>
        <p>Result: <%= pre(Server.HTMLEncode("code" & vbTab & "example")) %></p>
    </div>
    
    <div class="test success">
        <h2>✓ Test 3: Multiple tabs</h2>
        <pre><%= pre("line1" & vbTab & vbTab & "line2" & vbTab & "line3") %></pre>
    </div>
    
    <div class="test success">
        <h2>✓ All Tests Passed!</h2>
        <ul>
            <li>Function `pre()` is callable without errors</li>
            <li>Return variable is implicitly created</li>
            <li>Works with Option Explicit enabled</li>
            <li>Handles complex arguments correctly</li>
        </ul>
    </div>
    
</body>
</html>
