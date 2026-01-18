<%
Dim params
Set params = AxGetGet()
Response.Write "Type: " & TypeName(params) & vbCrLf
Response.Write "String: " & CStr(params) & vbCrLf
%>