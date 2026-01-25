<%@ Language="VBScript" %>
<html><body>
<%
On Error Resume Next

Set r = New RegExp
r.Pattern = "\d+"
Response.Write "OK - Pattern set"
%>
</body></html>
