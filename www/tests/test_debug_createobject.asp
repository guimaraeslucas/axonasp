<% @Language=VBScript %>
<%
Response.ContentType = "text/html"
On Error Resume Next
%>
<!DOCTYPE html>
<html>
<head>
    <title>Debug CreateObject</title>
</head>
<body>
    <h1>Debug CreateObject Tests</h1>
    
    <%
    ' Test 1: Create WScript.Shell
    Response.Write "<h2>Test 1: CreateObject('WScript.Shell')</h2>"
    Set objShell = CreateObject("WScript.Shell")
    Response.Write "<p>objShell = " & objShell & "</p>"
    Response.Write "<p>IsNull(objShell) = " & IsNull(objShell) & "</p>"
    Response.Write "<p>IsObject(objShell) = " & IsObject(objShell) & "</p>"
    Response.Write "<p>TypeName(objShell) = " & TypeName(objShell) & "</p>"
    Response.Write "<p>Err.Number = " & Err.Number & ", Err.Description = " & Err.Description & "</p>"
    Err.Clear()
    
    ' Test 2: Try calling a method
    If Not IsNull(objShell) Then
        Response.Write "<h2>Test 2: Call GetEnv</h2>"
        strPath = objShell.GetEnv("PATH")
        Response.Write "<p>PATH length = " & Len(strPath) & "</p>"
    Else
        Response.Write "<p>objShell is null, skipping method test</p>"
    End If
    
    %>
    
</body>
</html>
