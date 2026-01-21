<%@ Language=VBScript %>
<%
Response.Write "Starting tests<br>"

Dim obj
Response.Write "After Dim<br>"

Set obj = Nothing
Response.Write "After Set<br>"

Dim x
x = obj Is Nothing
Response.Write "x = obj Is Nothing: " & x & "<br>"

Dim y
y = (obj Is Nothing)
Response.Write "y = (obj Is Nothing): " & y & "<br>"

Response.Write "Done<br>"
%>
