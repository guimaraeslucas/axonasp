<%@ language="VBScript" %>
<% Option Explicit %>
<% debug_asp_code = "TRUE" %>

<h1>Simple colon test - DEBUG</h1>

<% 
Dim result
result = ""
If "hello" = "hello" Then result = "found" : End If
Response.Write "Result: " & result & "<br />"
%>
