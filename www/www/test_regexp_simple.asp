<%
Dim objRegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Pattern = "hello"

Response.Write "Pattern: " & objRegExp.Pattern & "<br>"
Response.Write "IgnoreCase: " & objRegExp.IgnoreCase & "<br>"
%>
