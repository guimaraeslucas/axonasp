<%
dim Bitmap(5, 5)
Bitmap(1, 1) = "01"
Response.Write "Bitmap(1,1) is " & Bitmap(1,1) & "<br>"
if Bitmap(1,1) = "01" then
    Response.Write "Match!<br>"
else
    Response.Write "No Match!<br>"
end if
%>
