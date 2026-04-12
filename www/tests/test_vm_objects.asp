<%
@ Language = "VBScript"
%>
<!DOCTYPE html>
<html>
<head><title>VM Objects Test</title></head>
<body>
<h1>AxonVM Objects Test</h1>

<%
Response.Write "<h2>Response Object</h2>"
Response.Write "<p>Response.Write works...</p>"
%>

<%
' Test Session object direct property setting
Session("TestKey1") = "TestValue1"
Response.Write "<h2>Session Object</h2>"
Response.Write "<p>Session.SessionID = " & Session.SessionID & "</p>"
Response.Write "<p>Session.Timeout = " & Session.Timeout & "</p>"
Response.Write "<p>Session.LCID = " & Session.LCID & "</p>"
Response.Write "<p>Session(""TestKey1"") = " & Session("TestKey1") & "</p>"

' Test Application object direct property setting
Application.Lock
Application("TestAppKey1") = "TestAppValue1"
Application.Unlock

Response.Write "<h2>Application Object</h2>"
Response.Write "<p>Application(""TestAppKey1"") = " & Application("TestAppKey1") & "</p>"
%>

<%
' Test Request object
Response.Write "<h2>Request Object</h2>"
Response.Write "<p>Request.ServerVariables(""SERVER_NAME"") = " & Request.ServerVariables("SERVER_NAME") & "</p>"
%>

<%
' Test Server object
Response.Write "<h2>Server Object</h2>"
Response.Write "<p>Server.ScriptTimeout = " & Server.ScriptTimeout & "</p>"
Response.Write "<p>Server.MapPath(""/"") = " & Server.MapPath("/") & "</p>"
%>

<%
' Test Err object
Response.Write "<h2>Err Object</h2>"
Response.Write "<p>Err.Number = " & Err.Number & "</p>"
Response.Write "<p>Err.Description = " & Err.Description & "</p>"
%>

</body>
</html>
