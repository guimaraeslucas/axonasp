<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <title>Operator Precedence Tests - Not vs Is</title>
    <style>
        .pass { color: green; }
        .fail { color: red; }
        .section { margin: 20px 0; padding: 10px; border: 1px solid #ccc; }
    </style>
</head>
<body>
    <h1>Operator Precedence Tests: Not vs Is</h1>
    <p>In VBScript, the Is operator has higher precedence than Not.</p>
    <p>Therefore, <code>Not X Is Nothing</code> is parsed as <code>Not (X Is Nothing)</code>, not <code>(Not X) Is Nothing</code>.</p>
    
<%
    Dim obj, result
    Dim passCount, failCount
    passCount = 0
    failCount = 0
    
    ' Test 1: Basic precedence with Nothing
    Response.Write "<div class='section'>"
    Response.Write "<h2>Test 1: Not obj Is Nothing (obj = Nothing)</h2>"
    Set obj = Nothing
    result = Not obj Is Nothing
    If result = False Then
        Response.Write "<p class='pass'>✓ PASS: Correctly evaluated as False</p>"
        passCount = passCount + 1
    Else
        Response.Write "<p class='fail'>✗ FAIL: Expected False, got " & result & "</p>"
        failCount = failCount + 1
    End If
    Response.Write "</div>"
    
    ' Test 2: Precedence with non-Nothing object
    Response.Write "<div class='section'>"
    Response.Write "<h2>Test 2: Not obj Is Nothing (obj <> Nothing)</h2>"
    Set obj = Server.CreateObject("Scripting.Dictionary")
    result = Not obj Is Nothing
    If result = True Then
        Response.Write "<p class='pass'>✓ PASS: Correctly evaluated as True</p>"
        passCount = passCount + 1
    Else
        Response.Write "<p class='fail'>✗ FAIL: Expected True, got " & result & "</p>"
        failCount = failCount + 1
    End If
    Response.Write "</div>"
    
    ' Test 3: Explicit parentheses should give same result
    Response.Write "<div class='section'>"
    Response.Write "<h2>Test 3: Not (obj Is Nothing) with explicit parens</h2>"
    result = Not (obj Is Nothing)
    If result = True Then
        Response.Write "<p class='pass'>✓ PASS: Correctly evaluated as True (same as without parens)</p>"
        passCount = passCount + 1
    Else
        Response.Write "<p class='fail'>✗ FAIL: Expected True, got " & result & "</p>"
        failCount = failCount + 1
    End If
    Response.Write "</div>"
    
    ' Test 4: In If statement with Nothing
    Response.Write "<div class='section'>"
    Response.Write "<h2>Test 4: If Not obj Is Nothing (obj = Nothing)</h2>"
    Set obj = Nothing
    If Not obj Is Nothing Then
        Response.Write "<p class='fail'>✗ FAIL: Should not enter this branch</p>"
        failCount = failCount + 1
    Else
        Response.Write "<p class='pass'>✓ PASS: Correctly took Else branch</p>"
        passCount = passCount + 1
    End If
    Response.Write "</div>"
    
    ' Test 5: In If statement with non-Nothing
    Response.Write "<div class='section'>"
    Response.Write "<h2>Test 5: If Not obj Is Nothing (obj <> Nothing)</h2>"
    Set obj = Server.CreateObject("Scripting.Dictionary")
    If Not obj Is Nothing Then
        Response.Write "<p class='pass'>✓ PASS: Correctly entered Then branch</p>"
        passCount = passCount + 1
    Else
        Response.Write "<p class='fail'>✗ FAIL: Should not enter Else branch</p>"
        failCount = failCount + 1
    End If
    Response.Write "</div>"
    
    ' Test 6: Combined expression with parentheses
    Response.Write "<div class='section'>"
    Response.Write "<h2>Test 6: Complex expression with And (using parentheses)</h2>"
    Set obj = Nothing
    Dim obj2
    Set obj2 = Server.CreateObject("Scripting.Dictionary")
    ' Explicitly use parentheses for clarity
    result = (Not obj Is Nothing) And (Not obj2 Is Nothing)
    If result = False Then
        Response.Write "<p class='pass'>✓ PASS: (Not (obj Is Nothing)) And (Not (obj2 Is Nothing)) = False And True = False</p>"
        passCount = passCount + 1
    Else
        Response.Write "<p class='fail'>✗ FAIL: Expected False, got " & result & "</p>"
        failCount = failCount + 1
    End If
    Response.Write "</div>"
    
    ' Summary
    Response.Write "<div class='section' style='background-color: " & IIf(failCount = 0, "#d4edda", "#f8d7da") & ";'>"
    Response.Write "<h2>Summary</h2>"
    Response.Write "<p><strong>Passed:</strong> " & passCount & "/" & (passCount + failCount) & "</p>"
    If failCount > 0 Then
        Response.Write "<p class='fail'><strong>Failed:</strong> " & failCount & "</p>"
    Else
        Response.Write "<p class='pass'><strong>All tests passed!</strong></p>"
    End If
    Response.Write "</div>"
%>
</body>
</html>
