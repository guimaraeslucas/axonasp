<%
Response.LCID = 1046
on error resume next

dim item
set item = new JSONpair

Response.Write "Line 1: Created JSONpair<br>"

item.name = "test"

if err.number <> 0 then
    Response.Write "Line 2: ERROR at step item.name assignment: " & err.description & " (Error " & err.number & ")<br>"
else
    Response.Write "Line 3: Successfully set item.name<br>"
end if

Response.Write "Line 4: item.name = " & item.name & "<br>"
%>
<!--#include file="json-teste/jsonObject.class.asp" -->
