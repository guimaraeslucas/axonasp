<%
dim rs : set rs = server.CreateObject("adodb.recordset")
Response.Write "TypeName of rs: " & typename(rs) & "<br>"
Response.Write "TypeName of rs.fields: " & typename(rs.fields) & "<br>"
%>
