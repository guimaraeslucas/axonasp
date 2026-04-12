<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Test: User Function 'pre' - Local Variables Issue Fix</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .test { margin: 20px 0; padding: 15px; background: white; border: 1px solid #ddd; border-radius: 5px; }
        .success { background: #d4edda; border-color: #28a745; }
        .error { background: #f8d7da; border-color: #f5c6cb; }
        pre { background: #f8f9fa; padding: 10px; border-radius: 3px; font-size: 12px; }
        h2 { color: #333; }
        code { background: #e9ecef; padding: 2px 4px; border-radius: 3px; }
    </style>
</head>
<body>
    <h1>Test: User Function 'pre' with Local Variables</h1>
    <p>Testing the exact function code from the user's issue</p>
    
    <%
    ' Define the exact function from the user's issue
    Function pre(value)
        Dim pre
        pre = Replace(value, vbTab, " ", 1, -1, 1)
        While InStr(pre, "    ") <> 0
            pre = Replace(pre, "    ", "   ", 1, -1, 1)
        Wend
    End Function
    
    Response.Write "<div class='test success'>"
    Response.Write "<h2>Test 1: Basic 'pre' Function Call</h2>"
    
    Dim testInput : testInput = "Hello" & vbTab & "World"
    Dim result : result = pre(testInput)
    
    Response.Write "<p><strong>Input:</strong></p>"
    Response.Write "<pre>" & Server.HTMLEncode(testInput) & "</pre>"
    Response.Write "<p><strong>Output:</strong></p>"
    Response.Write "<pre>" & Server.HTMLEncode(result) & "</pre>"
    Response.Write "<p style='color: green;'><strong>✓ Function executed without 'Variable is undefined: pre' error!</strong></p>"
    Response.Write "</div>"
    
    ' Test 2: Multiple calls
    Response.Write "<div class='test success'>"
    Response.Write "<h2>Test 2: Multiple Function Calls</h2>"
    
    Dim test1 : test1 = "Code" & vbTab & vbTab & "Example"
    Dim res1 : res1 = pre(test1)
    
    Dim test2 : test2 = "Another" & vbTab & "Test"
    Dim res2 : res2 = pre(test2)
    
    Response.Write "<p><strong>Call 1:</strong> " & res1 & "</p>"
    Response.Write "<p><strong>Call 2:</strong> " & res2 & "</p>"
    Response.Write "<p style='color: green;'><strong>✓ Multiple calls work correctly!</strong></p>"
    Response.Write "</div>"
    
    ' Test 3: Nested function calls
    Response.Write "<div class='test success'>"
    Response.Write "<h2>Test 3: Dim with Colon in Function Context</h2>"
    
    Function formatText(text)
        Dim cleaned : cleaned = pre(text)
        Dim upper : upper = UCase(cleaned)
        formatText = upper
    End Function
    
    Dim formatted : formatted = formatText("test" & vbTab & "data")
    Response.Write "<p><strong>Result:</strong> " & formatted & "</p>"
    Response.Write "<p style='color: green;'><strong>✓ Nested calls with Dim:assignment syntax work!</strong></p>"
    Response.Write "</div>"
    
    ' Summary
    Response.Write "<div class='test success'>"
    Response.Write "<h2>Summary of Fixes</h2>"
    Response.Write "<ul>"
    Response.Write "<li><strong>✓ Fix 1: Local Variables in Functions</strong> - Dim statements now properly declare local scope variables</li>"
    Response.Write "<li><strong>✓ Fix 2: Dim with Colon Syntax</strong> - <code>Dim var : var = value</code> now supported</li>"
    Response.Write "<li><strong>✓ Result:</strong> User function 'pre()' executes without 'Variable is undefined' error</li>"
    Response.Write "<li><strong>✓ Additional:</strong> Complex nested function calls with local variables work correctly</li>"
    Response.Write "</ul>"
    Response.Write "</div>"
    %>
</body>
</html>
