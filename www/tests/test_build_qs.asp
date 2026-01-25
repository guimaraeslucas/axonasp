<%
Dim dict
Set dict = CreateObject("Scripting.Dictionary")
dict.Add "name", "John Doe"
dict.Add "age", 25
dict.Add "city", "New York"

Dim qs
qs = AxBuildQueryString(dict)
Response.Write "Query String: " & qs & vbCrLf
%>