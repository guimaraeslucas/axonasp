<%
char = "["
if char = "{" then
  Response.Write "Block 1<br>"
elseif char = "[" then
  Response.Write "Block 2<br>"
else
  Response.Write "Block 3<br>"
end if
Response.Write "End<br>"
%>