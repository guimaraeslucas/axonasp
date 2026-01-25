<% 
' Test with ContentType
Response.ContentType = "text/html; charset=UTF-8"
result = 10 Mod 3
Response.Write "Result: " & result
%>