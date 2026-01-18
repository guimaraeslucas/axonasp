<%@ Page Language="VBScript" %>
<!DOCTYPE html>
<html>
<head>
    <title>Test Object Creation</title>
</head>
<body>
    <h1>Object Creation Test</h1>
    <%
    On Error Resume Next
    
    ' Test 1: Create a G3JSON object
    Response.Write "<h2>Test 1: G3JSON</h2>"
    Set obj = Server.CreateObject("G3JSON")
    If IsObject(obj) Then
        Response.Write "Object created successfully!<br>"
        Response.Write "Testing CallMethod...<br>"
        result = obj.NewObject()
        Response.Write "NewObject() returned: " & TypeName(result) & "<br>"
    Else
        Response.Write "Failed to create object<br>"
        If Err.Number <> 0 Then
            Response.Write "Error: " & Err.Description & "<br>"
        End If
    End If
    Set obj = Nothing
    Err.Clear
    
    ' Test 2: Create an ADO Connection
    Response.Write "<h2>Test 2: ADODB.Connection</h2>"
    Set conn = Server.CreateObject("ADODB.Connection")
    If IsObject(conn) Then
        Response.Write "Connection object created successfully!<br>"
        Response.Write "Connection State: " & conn.State & "<br>"
    Else
        Response.Write "Failed to create connection<br>"
        If Err.Number <> 0 Then
            Response.Write "Error: " & Err.Description & "<br>"
        End If
    End If
    Set conn = Nothing
    Err.Clear
    %>
</body>
</html>
