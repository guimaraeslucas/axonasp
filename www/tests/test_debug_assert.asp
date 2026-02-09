<% @Language=VBScript %>
<%
Response.ContentType = "text/html"
On Error Resume Next
%>
<!DOCTYPE html>
<html>
<head>
    <title>Debug Assert</title>
</head>
<body>
    <h1>Debug AssertNotNull</h1>
    
    <%
    Dim testsPassed, testsFailed, testResults
    testsPassed = 0
    testsFailed = 0
    testResults = ""
    
    Function AssertNotNull(obj, testName)
        Response.Write "<p>DEBUG: In AssertNotNull, obj=" & obj & ", IsNothing=" & (obj Is Nothing) & ", NotNothing=" & Not (obj Is Nothing) & "</p>"
        If Not (obj Is Nothing) Then
            testsPassed = testsPassed + 1
            testResults = testResults & "<div style='color: green;'>[PASS] " & testName & "</div>"
            Response.Write "<p>DEBUG: Incrementing testsPassed to " & testsPassed & "</p>"
        Else
            testsFailed = testsFailed + 1
            testResults = testResults & "<div style='color: red;'>[FAIL] " & testName & " (expected non-null object)</div>"
            Response.Write "<p>DEBUG: Incrementing testsFailed to " & testsFailed & "</p>"
        End If
    End Function
    
    Set objShell = CreateObject("WScript.Shell")
    Response.Write "<p>objShell = " & objShell & "</p>"
    
    AssertNotNull objShell, "Test CreateObject"
    
    Response.Write "<p>testsPassed = " & testsPassed & ", testsFailed = " & testsFailed & "</p>"
    
    %>
    
</body>
</html>
