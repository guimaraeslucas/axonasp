<%@ Language="VBScript" %>
<html><body>
<%
On Error Resume Next

Set r = New RegExp
r.Pattern = "\d+"
result = r.Test("abc123def")
Response.Write "Test result: " & result
%>
</body></html>
