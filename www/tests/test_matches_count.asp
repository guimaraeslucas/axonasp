<%@ Language="VBScript" %>
<html><body>
<%
On Error Resume Next

Set r = New RegExp
r.Pattern = "\d+"
r.Global = True
Set matches = r.Execute("abc123def456")
Response.Write "Matches count: " & matches.Count
%>
</body></html>
