<%
Response.Write "db type: " & typename(db) & "<br>"
Response.Write "db.rs call...<br>"
dim rs : set rs = db.rs()
Response.Write "rs type: " & typename(rs) & "<br>"
%>
