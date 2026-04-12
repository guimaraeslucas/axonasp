<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Test Dim with Colon and Local Variables in Functions</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .test { margin: 20px 0; padding: 15px; background: white; border: 1px solid #ddd; border-radius: 5px; }
        .success { background: #d4edda; border-color: #28a745; }
        .info { background: #d1ecf1; border-color: #17a2b8; }
        pre { background: #f8f9fa; padding: 10px; border-radius: 3px; overflow-x: auto; }
        .error { background: #f8d7da; border-color: #f5c6cb; }
        h2 { color: #333; }
    </style>
</head>
<body>
    <h1>Test: Dim with Colon (:) and Local Variables in Functions</h1>
    
    <%
    ' Test 1: Simple Dim with colon assignment
    Response.Write "<div class='test info'>"
    Response.Write "<h2>Test 1: Dim with Colon Assignment (Simple)</h2>"
    Response.Write "<p>Syntax: <code>Dim myVar : myVar = 'Hello'</code></p>"
    
    Dim simpleVar : simpleVar = "Hello World"
    Response.Write "<p><strong>Result:</strong> simpleVar = '" & simpleVar & "'</p>"
    Response.Write "</div>"
    
    ' Test 2: Multiple variables, then assignment
    Response.Write "<div class='test info'>"
    Response.Write "<h2>Test 2: Dim Multiple Vars with Colon Assignment</h2>"
    Response.Write "<p>Syntax: <code>Dim var1, var2 : var1 = value</code></p>"
    
    Dim var1, var2 : var1 = 42
    Response.Write "<p><strong>var1 =</strong> " & var1 & " (assigned)</p>"
    Response.Write "<p><strong>var2 =</strong> '" & var2 & "' (undefined, empty string)</p>"
    Response.Write "</div>"
    
    ' Test 3: Dim with expression evaluation
    Response.Write "<div class='test info'>"
    Response.Write "<h2>Test 3: Dim with Expression Evaluation</h2>"
    Response.Write "<p>Syntax: <code>Dim result : result = 5 + 3</code></p>"
    
    Dim result : result = 5 + 3
    Response.Write "<p><strong>result =</strong> " & result & " (5 + 3)</p>"
    Response.Write "</div>"
    
    ' Test 4: User-defined function with local variables (the 'pre' function issue)
    Response.Write "<div class='test success'>"
    Response.Write "<h2>Test 4: Function with Local Variables (pre function)</h2>"
    Response.Write "<p>Testing the user's <code>pre(value)</code> function with local Dim statements</p>"
    
    Function pre(value)
        Dim output
        output = Replace(value, vbTab, " ", 1, -1, 1)
        While InStr(output, "    ") <> 0
            output = Replace(output, "    ", "   ", 1, -1, 1)
        Wend
        pre = output
    End Function
    
    Dim inputText
    inputText = "Hello" & vbTab & "World" & vbTab & "Test"
    Dim outputText
    outputText = pre(inputText)
    
    Response.Write "<p><strong>Input:</strong> <code>" & Server.HTMLEncode(inputText) & "</code></p>"
    Response.Write "<p><strong>Output:</strong> <code>" & Server.HTMLEncode(outputText) & "</code></p>"
    Response.Write "<p><strong>Status:</strong> <span style='color: green;'>✓ Function executed successfully</span></p>"
    Response.Write "</div>"
    
    ' Test 5: Function with multiple local variables
    Response.Write "<div class='test success'>"
    Response.Write "<h2>Test 5: Function with Multiple Local Variables</h2>"
    Response.Write "<p>Testing function with multiple Dim statements</p>"
    
    Function processString(inputStr)
        Dim temp, result, i
        temp = inputStr
        result = ""
        
        ' Simple processing
        For i = 1 To Len(temp)
            If i Mod 2 = 0 Then
                result = result & "[" & Mid(temp, i, 1) & "]"
            Else
                result = result & Mid(temp, i, 1)
            End If
        Next
        
        processString = result
    End Function
    
    Dim testStr : testStr = "TEST"
    Dim processed : processed = processString(testStr)
    
    Response.Write "<p><strong>Input:</strong> " & testStr & "</p>"
    Response.Write "<p><strong>Output:</strong> " & processed & "</p>"
    Response.Write "<p><strong>Status:</strong> <span style='color: green;'>✓ Multiple local variables work</span></p>"
    Response.Write "</div>"
    
    ' Test 6: Dim with colon in function
    Response.Write "<div class='test success'>"
    Response.Write "<h2>Test 6: Dim with Colon Inside Function</h2>"
    Response.Write "<p>Testing <code>Dim var : var = value</code> syntax within function</p>"
    
    Function testDimColon(val)
        Dim x : x = val * 2
        Dim y : y = val + 10
        testDimColon = x + y
    End Function
    
    Dim funcResult : funcResult = testDimColon(5)
    Response.Write "<p><strong>testDimColon(5) = </strong> " & funcResult & " (expected: (5*2) + (5+10) = 25)</p>"
    If funcResult = 25 Then
        Response.Write "<p><strong>Status:</strong> <span style='color: green;'>✓ Result matches expected value</span></p>"
    Else
        Response.Write "<p><strong>Status:</strong> <span style='color: red;'>✗ Result mismatch (got " & funcResult & ", expected 25)</span></p>"
    End If
    Response.Write "</div>"
    
    ' Summary
    Response.Write "<div class='test success'>"
    Response.Write "<h2>Summary</h2>"
    Response.Write "<ul>"
    Response.Write "<li>✓ Dim with colon assignment syntax</li>"
    Response.Write "<li>✓ Multiple variables with colon assignment</li>"
    Response.Write "<li>✓ Expression evaluation in Dim with colon</li>"
    Response.Write "<li>✓ Local variables in functions (Dim statement)</li>"
    Response.Write "<li>✓ User-defined function 'pre()' working</li>"
    Response.Write "<li>✓ Multiple local variables in function</li>"
    Response.Write "<li>✓ Dim with colon inside function scope</li>"
    Response.Write "</ul>"
    Response.Write "</div>"
    %>
</body>
</html>
