<%@ Language="VBScript" %>
<html>
<head><title>RegExp Final Test</title>
<style>
body { font-family: Arial; margin: 20px; }
.test { padding: 10px; margin: 10px 0; border: 1px solid #ccc; }
.pass { background-color: #e8f5e9; color: green; }
.fail { background-color: #ffebee; color: red; }
h1 { color: #333; }
h2 { color: #666; margin-top: 20px; }
</style>
</head>
<body>

<h1>RegExp Implementation Final Test</h1>
<p>Testing if RegExp object works correctly after fixes.</p>

<%
    ' Keep track of test results
    passCount = 0
    failCount = 0
    
    ' ===== TEST 1: Create object
    response.write "<h2>TEST 1: Object Creation</h2>"
    response.write "<div class='test'>"
    
    On Error Resume Next
    Set regex = Server.CreateObject("G3REGEXP")
    
    if Err.Number <> 0 then
        response.write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        failCount = failCount + 1
    else
        response.write "<div class='pass'>PASS: G3REGEXP object created</div>"
        passCount = passCount + 1
    end if
    
    response.write "</div>"
    
    ' ===== TEST 2: Set Pattern via direct assignment
    response.write "<h2>TEST 2: Set Pattern (assignment)</h2>"
    response.write "<div class='test'>"
    
    On Error Resume Next
    Err.Clear
    regex.Pattern = "\d+"
    
    if Err.Number <> 0 then
        response.write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        failCount = failCount + 1
    else
        response.write "<div class='pass'>PASS: Pattern assigned</div>"
        passCount = passCount + 1
    end if
    
    response.write "</div>"
    
    ' ===== TEST 3: Get Pattern
    response.write "<h2>TEST 3: Get Pattern</h2>"
    response.write "<div class='test'>"
    
    On Error Resume Next
    Err.Clear
    pat = regex.GetProperty("Pattern")
    
    if Err.Number <> 0 then
        response.write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        failCount = failCount + 1
    elseif pat = "\d+" then
        response.write "<div class='pass'>PASS: Pattern retrieved correctly: '" & pat & "'</div>"
        passCount = passCount + 1
    else
        response.write "<div class='fail'>FAIL: Pattern value incorrect. Expected '\d+', got '" & pat & "'</div>"
        failCount = failCount + 1
    end if
    
    response.write "</div>"
    
    ' ===== TEST 4: Test() method - should match
    response.write "<h2>TEST 4: Test() - Should Match</h2>"
    response.write "<div class='test'>"
    response.write "Input: 'abc123def', Pattern: '\d+'" & "<br>"
    
    On Error Resume Next
    Err.Clear
    result = regex.Test("abc123def")
    
    if Err.Number <> 0 then
        response.write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        failCount = failCount + 1
    elseif result = True then
        response.write "<div class='pass'>PASS: Test() returned TRUE</div>"
        passCount = passCount + 1
    else
        response.write "<div class='fail'>FAIL: Test() returned FALSE, expected TRUE</div>"
        failCount = failCount + 1
    end if
    
    response.write "</div>"
    
    ' ===== TEST 5: Test() method - should NOT match
    response.write "<h2>TEST 5: Test() - Should NOT Match</h2>"
    response.write "<div class='test'>"
    response.write "Input: 'abcdef', Pattern: '\d+'" & "<br>"
    
    On Error Resume Next
    Err.Clear
    result = regex.Test("abcdef")
    
    if Err.Number <> 0 then
        response.write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        failCount = failCount + 1
    elseif result = False then
        response.write "<div class='pass'>PASS: Test() returned FALSE</div>"
        passCount = passCount + 1
    else
        response.write "<div class='fail'>FAIL: Test() returned TRUE, expected FALSE</div>"
        failCount = failCount + 1
    end if
    
    response.write "</div>"
    
    ' ===== TEST 6: Set Global flag and IgnoreCase
    response.write "<h2>TEST 6: Set Properties (Global, IgnoreCase)</h2>"
    response.write "<div class='test'>"
    
    On Error Resume Next
    Err.Clear
    regex.Global = True
    regex.IgnoreCase = True
    
    if Err.Number <> 0 then
        response.write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        failCount = failCount + 1
    else
        response.write "<div class='pass'>PASS: Properties set</div>"
        passCount = passCount + 1
    end if
    
    response.write "</div>"
    
    ' ===== TEST 7: Execute() method
    response.write "<h2>TEST 7: Execute() Method</h2>"
    response.write "<div class='test'>"
    response.write "Input: 'abc123def456', Pattern: '\d+', Global: TRUE" & "<br>"
    
    On Error Resume Next
    Err.Clear
    Set matches = regex.Execute("abc123def456")
    
    if Err.Number <> 0 then
        response.write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        failCount = failCount + 1
    elseif matches Is Nothing then
        response.write "<div class='fail'>FAIL: Execute() returned Nothing</div>"
        failCount = failCount + 1
    else
        response.write "Matches object type: " & TypeName(matches) & "<br>"
        count = matches.Count
        response.write "Match count: " & count & "<br>"
        
        if count = 2 then
            response.write "<div class='pass'>PASS: Found 2 matches as expected</div>"
            passCount = passCount + 1
            
            ' Show the matches
            For i = 0 To count - 1
                Set match = matches.Item(i)
                response.write "Match " & (i+1) & ": Value='" & match.Value & "', Index=" & match.Index & "<br>"
            Next
        else
            response.write "<div class='fail'>FAIL: Expected 2 matches, got " & count & "</div>"
            failCount = failCount + 1
        end if
    end if
    
    response.write "</div>"
    
    ' ===== TEST 8: Replace() method
    response.write "<h2>TEST 8: Replace() Method</h2>"
    response.write "<div class='test'>"
    response.write "Input: 'abc123def456', Pattern: '\d+', Replacement: 'X'" & "<br>"
    
    On Error Resume Next
    Err.Clear
    replaced = regex.Replace("abc123def456", "X")
    
    if Err.Number <> 0 then
        response.write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        failCount = failCount + 1
    else
        response.write "Result: '" & replaced & "'<br>"
        if replaced = "abcXdefX" then
            response.write "<div class='pass'>PASS: Replace() worked correctly</div>"
            passCount = passCount + 1
        else
            response.write "<div class='fail'>FAIL: Expected 'abcXdefX', got '" & replaced & "'</div>"
            failCount = failCount + 1
        end if
    end if
    
    response.write "</div>"
    
    ' Summary
    response.write "<h2>SUMMARY</h2>"
    response.write "<div style='font-size: 18px; font-weight: bold;'>"
    response.write "Passed: <span style='color: green;'>" & passCount & "</span>, "
    response.write "Failed: <span style='color: red;'>" & failCount & "</span>"
    response.write "</div>"
%>

</body>
</html>
