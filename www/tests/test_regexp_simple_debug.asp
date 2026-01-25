<%@ Language="VBScript" %>
<!--
Simple Debug Test for RegExp
Tests object creation, property setting, and method calls
-->
<html>
<head>
    <title>RegExp Simple Debug Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .test { margin: 10px 0; padding: 10px; border: 1px solid #ccc; }
        .pass { color: green; background-color: #e8f5e9; }
        .fail { color: red; background-color: #ffebee; }
        .info { color: blue; }
        h2 { color: #333; }
        code { background-color: #f5f5f5; padding: 2px 5px; }
    </style>
</head>
<body>

<h1>RegExp - Simple Debug Test</h1>
<p>Testing object creation, properties, and basic methods.</p>

<%
    ' Test 1: Create object
    response.write("<div class='test'>" & vbCrLf)
    response.write("<strong>Test 1: Object Creation</strong><br>" & vbCrLf)
    
    On Error Resume Next
    Set regex = Server.CreateObject("G3REGEXP")
    If Err.Number <> 0 Then
        response.write("<span class='fail'>FAIL: Could not create object. Error: " & Err.Description & "</span><br>" & vbCrLf)
    Else
        response.write("<span class='pass'>PASS: Object created</span><br>" & vbCrLf)
        response.write("Object type: " & TypeName(regex) & "<br>" & vbCrLf)
    End If
    response.write("</div>" & vbCrLf)
    
    ' Test 2: Check if object has methods
    response.write("<div class='test'>" & vbCrLf)
    response.write("<strong>Test 2: Check Method Existence (indirect check)</strong><br>" & vbCrLf)
    response.write("We'll test by calling them...<br>" & vbCrLf)
    response.write("</div>" & vbCrLf)
    
    ' Test 3: Set Pattern directly via property
    response.write("<div class='test'>" & vbCrLf)
    response.write("<strong>Test 3: Set Pattern (via method call)</strong><br>" & vbCrLf)
    
    On Error Resume Next
    Err.Clear
    ' Try setting pattern by calling it as a method
    regex.Pattern = "\d+"
    
    If Err.Number <> 0 Then
        response.write("<span class='info'>Note: Direct property assignment failed, trying SetProperty method...</span><br>" & vbCrLf)
        Err.Clear
        regex.SetProperty "Pattern", "\d+"
        If Err.Number <> 0 Then
            response.write("<span class='fail'>FAIL: SetProperty also failed. Error: " & Err.Description & "</span><br>" & vbCrLf)
        Else
            response.write("<span class='pass'>PASS: SetProperty worked</span><br>" & vbCrLf)
        End If
    Else
        response.write("<span class='pass'>PASS: Pattern property set to '\d+'</span><br>" & vbCrLf)
    End If
    response.write("</div>" & vbCrLf)
    
    ' Test 4: Get Pattern
    response.write("<div class='test'>" & vbCrLf)
    response.write("<strong>Test 4: Get Pattern</strong><br>" & vbCrLf)
    
    On Error Resume Next
    Err.Clear
    patternValue = regex.GetProperty("Pattern")
    
    If Err.Number <> 0 Then
        response.write("<span class='fail'>FAIL: GetProperty failed. Error: " & Err.Description & "</span><br>" & vbCrLf)
    Else
        response.write("Pattern value: <code>" & patternValue & "</code><br>" & vbCrLf)
        if patternValue = "\d+" then
            response.write("<span class='pass'>PASS: Pattern correctly stored and retrieved</span><br>" & vbCrLf)
        else
            response.write("<span class='fail'>FAIL: Pattern value mismatch. Expected '\d+', got '" & patternValue & "'</span><br>" & vbCrLf)
        end if
    End If
    response.write("</div>" & vbCrLf)
    
    ' Test 5: Test() method with valid pattern
    response.write("<div class='test'>" & vbCrLf)
    response.write("<strong>Test 5: Test() Method - Should Match</strong><br>" & vbCrLf)
    response.write("Pattern: <code>\d+</code>, Input: <code>abc123def</code><br>" & vbCrLf)
    
    On Error Resume Next
    Err.Clear
    result = regex.Test("abc123def")
    
    If Err.Number <> 0 Then
        response.write("<span class='fail'>FAIL: Test() failed. Error: " & Err.Description & "</span><br>" & vbCrLf)
    Else
        if result Then
            response.write("<span class='pass'>PASS: Test() returned TRUE (match found)</span><br>" & vbCrLf)
        else
            response.write("<span class='fail'>FAIL: Test() returned FALSE (no match found)</span><br>" & vbCrLf)
        end if
        response.write("Result value: " & result & " (type: " & TypeName(result) & ")<br>" & vbCrLf)
    End If
    response.write("</div>" & vbCrLf)
    
    ' Test 6: Test() method with non-matching input
    response.write("<div class='test'>" & vbCrLf)
    response.write("<strong>Test 6: Test() Method - Should NOT Match</strong><br>" & vbCrLf)
    response.write("Pattern: <code>\d+</code>, Input: <code>abcdef</code><br>" & vbCrLf)
    
    On Error Resume Next
    Err.Clear
    result = regex.Test("abcdef")
    
    If Err.Number <> 0 Then
        response.write("<span class='fail'>FAIL: Test() failed. Error: " & Err.Description & "</span><br>" & vbCrLf)
    Else
        if Not result Then
            response.write("<span class='pass'>PASS: Test() returned FALSE (no match)</span><br>" & vbCrLf)
        else
            response.write("<span class='fail'>FAIL: Test() returned TRUE (unexpected match)</span><br>" & vbCrLf)
        end if
    End If
    response.write("</div>" & vbCrLf)
    
    ' Test 7: Execute() method
    response.write("<div class='test'>" & vbCrLf)
    response.write("<strong>Test 7: Execute() Method</strong><br>" & vbCrLf)
    response.write("Pattern: <code>\d+</code>, Input: <code>abc123def456</code><br>" & vbCrLf)
    
    On Error Resume Next
    Err.Clear
    Set matches = regex.Execute("abc123def456")
    
    If Err.Number <> 0 Then
        response.write("<span class='fail'>FAIL: Execute() failed. Error: " & Err.Description & "</span><br>" & vbCrLf)
    Else
        response.write("Execute() returned: " & TypeName(matches) & "<br>" & vbCrLf)
        if Not (matches Is Nothing) Then
            response.write("Matches count: " & matches.Count & "<br>" & vbCrLf)
            if matches.Count > 0 Then
                response.write("<span class='pass'>PASS: Execute() found matches</span><br>" & vbCrLf)
                response.write("First match: " & matches.Item(0).Value & " at index " & matches.Item(0).Index & "<br>" & vbCrLf)
            else
                response.write("<span class='fail'>FAIL: Execute() found no matches</span><br>" & vbCrLf)
            end if
        else
            response.write("<span class='fail'>FAIL: Execute() returned Nothing</span><br>" & vbCrLf)
        end if
    End If
    response.write("</div>" & vbCrLf)
    
    ' Test 8: Replace() method
    response.write("<div class='test'>" & vbCrLf)
    response.write("<strong>Test 8: Replace() Method</strong><br>" & vbCrLf)
    response.write("Pattern: <code>\d+</code>, Input: <code>abc123def456</code>, Replacement: <code>X</code><br>" & vbCrLf)
    
    On Error Resume Next
    Err.Clear
    replaced = regex.Replace("abc123def456", "X")
    
    If Err.Number <> 0 Then
        response.write("<span class='fail'>FAIL: Replace() failed. Error: " & Err.Description & "</span><br>" & vbCrLf)
    Else
        response.write("Result: <code>" & replaced & "</code><br>" & vbCrLf)
        if replaced = "abcXdefX" Then
            response.write("<span class='pass'>PASS: Replace() worked correctly</span><br>" & vbCrLf)
        else
            response.write("<span class='fail'>FAIL: Replace() result incorrect. Expected 'abcXdefX', got '" & replaced & "'</span><br>" & vbCrLf)
        end if
    End If
    response.write("</div>" & vbCrLf)

%>

</body>
</html>
