<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <title>AxInclude Debug Test</title>
</head>
<body>
    <h1>AxInclude Debug Test</h1>
    
    <h2>Test 1: Direct Assignment</h2>
    <%
        Dim result
        result = AxInclude("config.inc")
        Response.Write "Direct result = " & result & "<br>" & vbCrLf
        Response.Write "Result type check: " & TypeName(result) & "<br>" & vbCrLf
        Response.Write "Result = True: " & (result = True) & "<br>" & vbCrLf
        Response.Write "Result = False: " & (result = False) & "<br>" & vbCrLf
        Response.Write "Len(result): " & Len(result) & "<br>" & vbCrLf
    %>
    
    <h2>Test 2: If Statement</h2>
    <%
        Response.Write "Before If<br>" & vbCrLf
        If AxInclude("config.inc") Then
            Response.Write "AxInclude returned TRUE<br>" & vbCrLf
        Else
            Response.Write "AxInclude returned FALSE<br>" & vbCrLf
        End If
        Response.Write "After If<br>" & vbCrLf
    %>

    <h2>Test 3: Direct Boolean Check</h2>
    <%
        If AxInclude("config.inc") = True Then
            Response.Write "Explicitly = True<br>" & vbCrLf
        End If
    %>
</body>
</html>
