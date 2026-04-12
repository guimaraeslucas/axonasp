<%
dim item
set item = new JSONpair
Response.Write "Created JSONpair<br>"
Response.Write "TypeName: " & TypeName(item) & "<br>"
Response.Write "IsObject: " & IsObject(item) & "<br>"

on error resume next
item.name = "test"
if err.number <> 0 then
    Response.Write "ERROR setting name: " & err.description & " (" & err.number & ")<br>"
else
    Response.Write "Successfully set name<br>"
end if
%>
<!--#include file="json-teste/jsonObject.class.asp" -->
