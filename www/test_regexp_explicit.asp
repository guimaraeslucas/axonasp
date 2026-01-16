<%
Option Explicit
Dim objRegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
Response.Write "Success: " & objRegExp.IgnoreCase & "<br>"
%>
