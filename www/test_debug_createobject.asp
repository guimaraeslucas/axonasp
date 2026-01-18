<%@ Language="VBScript" %>
<html>
<head><title>Debug CreateObject</title></head>
<body>
<h1>Debug CreateObject</h1>

<%
Response.Write "<p>Step 1: Check if Server object exists</p>"
If IsObject(Server) Then
    Response.Write "<p style='color:green'>Server object exists</p>"
Else
    Response.Write "<p style='color:red'>Server object does NOT exist</p>"
End If

Response.Write "<p>Step 2: Try to create G3JSON (known working)</p>"
On Error Resume Next
Dim objJSON
Set objJSON = Server.CreateObject("G3JSON")
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Error: " & Err.Number & " - " & Err.Description & " - " & Err.Source & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>G3JSON created successfully</p>"
    If IsObject(objJSON) Then
        Response.Write "<p style='color:green'>objJSON is an object</p>"
    End If
End If

Response.Write "<p>Step 3: Try to create G3TEMPLATE</p>"
Dim objTemplate
Set objTemplate = Server.CreateObject("G3TEMPLATE")
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Error: " & Err.Number & " - " & Err.Description & " - " & Err.Source & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>G3TEMPLATE created successfully</p>"
    If IsObject(objTemplate) Then
        Response.Write "<p style='color:green'>objTemplate is an object</p>"
    End If
End If

Response.Write "<p>Step 4: Try to create G3MAIL</p>"
Dim objMail
Set objMail = Server.CreateObject("G3MAIL")
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Error: " & Err.Number & " - " & Err.Description & " - " & Err.Source & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>G3MAIL created successfully</p>"
    If IsObject(objMail) Then
        Response.Write "<p style='color:green'>objMail is an object</p>"
    End If
End If
%>

</body>
</html>
