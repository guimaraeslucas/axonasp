<%@ language="VBScript" %>
<% Option Explicit %>

<h1>Simple colon test</h1>

<% 
Dim result
result = ""
If "hello" = "hello" Then result = "found" : End If
Response.Write "Result: " & result & "<br />"
%>
