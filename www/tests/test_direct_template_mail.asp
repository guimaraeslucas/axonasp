<%@ Language="VBScript" %>
<html>
<head><title>Direct Test</title></head>
<body>
<h1>Direct CreateObject Test (No Error Handling)</h1>

<h2>Test G3TEMPLATE</h2>
<%
Dim objTemplate
Set objTemplate = Server.CreateObject("G3TEMPLATE")
Response.Write "<p>Object type: " & TypeName(objTemplate) & "</p>"
Response.Write "<p>IsObject: " & IsObject(objTemplate) & "</p>"
%>

<h2>Test G3MAIL</h2>
<%
Dim objMail
Set objMail = Server.CreateObject("G3MAIL")
Response.Write "<p>Object type: " & TypeName(objMail) & "</p>"
Response.Write "<p>IsObject: " & IsObject(objMail) & "</p>"
%>

</body>
</html>
