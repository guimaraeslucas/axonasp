<% @Language=VBScript %>
<%
Response.ContentType = "text/html"
On Error Resume Next

Dim testsPassed, testsFailed, testResults
testsPassed = 0
testsFailed = 0
testResults = ""

Function AssertNotNull(obj, testName)
    boolResult = Not (obj Is Nothing)
    Response.Write "<p>DEBUG: testName='" & testName & "', obj=" & obj & ", boolResult=" & boolResult & ", isTruthy=" & (boolResult <> 0) & "</p>"
    
    If Not (obj Is Nothing) Then
        testsPassed = testsPassed + 1
        testResults = testResults & "<div style='color: green;'>[PASS] " & testName & "</div>"
        Response.Write "<p>PASS incremented testsPassed to " & testsPassed & "</p>"
    Else
        testsFailed = testsFailed + 1
        testResults = testResults & "<div style='color: red;'>[FAIL] " & testName & "</div>"
        Response.Write "<p>FAIL incremented testsFailed to " & testsFailed & "</p>"
    End If
End Function

%>
<!DOCTYPE html>
<html>
<head>
    <title>Debug Assert Function</title>
</head>
<body>
    <h1>Debug Assert Function</h1>
    
    <%
    Set objShell = CreateObject("WScript.Shell")
    Response.Write "<p>Created objShell = " & objShell & "</p>"
    
    AssertNotNull objShell, "CreateObject('WScript.Shell')"
    
    Response.Write "<h2>Final Results</h2>"
    Response.Write "<p>testsPassed = " & testsPassed & "</p>"
    Response.Write "<p>testsFailed = " & testsFailed & "</p>"
    Response.Write "<div>" & testResults & "</div>"
    
    %>
    
</body>
</html>
