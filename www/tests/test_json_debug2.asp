<%
dim rs : set rs=db.rs
Response.Write "TypeName of rs (before open): " & typename(rs) & "<br>"
Response.Write "TypeName of rs.fields: " & typename(rs.fields) & "<br>"
rs.open("select top 5 * from testdata")
Response.Write "TypeName of rs (after open): " & typename(rs) & "<br>"
Response.Write "Fields count: " & rs.fields.count & "<br>"

for i = 0 to rs.fields.count - 1
    Response.Write "Field " & i & " name: " & rs.fields(i).name & "<br>"
next
%>
