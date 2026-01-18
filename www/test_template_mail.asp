<%@ Language="VBScript" %>
<html>
<head><title>Template and Mail Test</title></head>
<body>
<h1>Template and Mail Library Test</h1>

<h2>Testing G3TEMPLATE</h2>
<%
On Error Resume Next
Dim objTemplate
Set objTemplate = Server.CreateObject("G3TEMPLATE")
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR creating G3TEMPLATE: " & Err.Description & "</p>"
    Err.Clear
ElseIf IsObject(objTemplate) Then
    Response.Write "<p style='color:green'>SUCCESS: G3TEMPLATE object created</p>"
Else
    Response.Write "<p style='color:red'>ERROR: G3TEMPLATE returned but not an object</p>"
End If
Set objTemplate = Nothing
%>

<h2>Testing TEMPLATE (alias)</h2>
<%
Set objTemplate = Server.CreateObject("TEMPLATE")
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR creating TEMPLATE: " & Err.Description & "</p>"
    Err.Clear
ElseIf IsObject(objTemplate) Then
    Response.Write "<p style='color:green'>SUCCESS: TEMPLATE object created</p>"
Else
    Response.Write "<p style='color:red'>ERROR: TEMPLATE returned but not an object</p>"
End If
Set objTemplate = Nothing
%>

<h2>Testing G3MAIL</h2>
<%
Dim objMail
Set objMail = Server.CreateObject("G3MAIL")
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR creating G3MAIL: " & Err.Description & "</p>"
    Err.Clear
ElseIf IsObject(objMail) Then
    Response.Write "<p style='color:green'>SUCCESS: G3MAIL object created</p>"
Else
    Response.Write "<p style='color:red'>ERROR: G3MAIL returned but not an object</p>"
End If
Set objMail = Nothing
%>

<h2>Testing MAIL (alias)</h2>
<%
Set objMail = Server.CreateObject("MAIL")
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR creating MAIL: " & Err.Description & "</p>"
    Err.Clear
ElseIf IsObject(objMail) Then
    Response.Write "<p style='color:green'>SUCCESS: MAIL object created</p>"
Else
    Response.Write "<p style='color:red'>ERROR: MAIL returned but not an object</p>"
End If
Set objMail = Nothing
%>

</body>
</html>
