<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <title>Not/Is Precedence Test</title>
</head>
<body>
    <h1>Not/Is Precedence Test</h1>
    
<%
    Dim obj
    Set obj = Nothing
    
    Response.Write "<h2>Test 1: Not obj Is Nothing (should be False when obj=Nothing)</h2>"
    If Not obj Is Nothing Then
        Response.Write "<p style='color: red;'>✗ FAIL: Not obj Is Nothing evaluated as True (incorrect)</p>"
    Else
        Response.Write "<p style='color: green;'>✓ PASS: Not obj Is Nothing evaluated as False (correct)</p>"
    End If
    
    Response.Write "<h2>Test 2: obj Is Nothing (should be True)</h2>"
    If obj Is Nothing Then
        Response.Write "<p style='color: green;'>✓ PASS: obj Is Nothing is True</p>"
    Else
        Response.Write "<p style='color: red;'>✗ FAIL: obj Is Nothing is False</p>"
    End If
    
    Set obj = Server.CreateObject("Scripting.Dictionary")
    
    Response.Write "<h2>Test 3: Not obj Is Nothing (should be True when obj<>Nothing)</h2>"
    If Not obj Is Nothing Then
        Response.Write "<p style='color: green;'>✓ PASS: Not obj Is Nothing evaluated as True (correct)</p>"
    Else
        Response.Write "<p style='color: red;'>✗ FAIL: Not obj Is Nothing evaluated as False (incorrect)</p>"
    End If
    
    Response.Write "<h2>Test 4: obj Is Nothing (should be False)</h2>"
    If obj Is Nothing Then
        Response.Write "<p style='color: red;'>✗ FAIL: obj Is Nothing is True</p>"
    Else
        Response.Write "<p style='color: green;'>✓ PASS: obj Is Nothing is False</p>"
    End If
    
    Set obj = Nothing
    Response.Write "<h2>Test 5: Explicit parentheses - Not (obj Is Nothing) when obj=Nothing</h2>"
    If Not (obj Is Nothing) Then
        Response.Write "<p style='color: red;'>✗ FAIL: Not (obj Is Nothing) is True</p>"
    Else
        Response.Write "<p style='color: green;'>✓ PASS: Not (obj Is Nothing) is False</p>"
    End If
%>
</body>
</html>
