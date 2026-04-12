<%
Response.Write "Starting test<br>"

Set objRegExp = New RegExp
Response.Write "Created RegExp object<br>"

objRegExp.IgnoreCase = True
Response.Write "Set IgnoreCase<br>"

objRegExp.Pattern = "hello"
Response.Write "Set Pattern<br>"

Response.Write "Pattern: " & objRegExp.Pattern & "<br>"
Response.Write "IgnoreCase: " & objRegExp.IgnoreCase & "<br>"

Response.Write "Test completed!<br>"
%>
