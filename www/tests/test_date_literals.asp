<%
Dim d1, d2
d1 = #1/19/2026#
d2 = #2026-01-19#

Response.Write "Date 1: " & d1 & "<br>"
Response.Write "Date 2: " & d2 & "<br>"
Response.Write "TypeName: " & TypeName(d1) & "<br>"

If IsDate(d1) Then
    Response.Write "d1 is a valid date.<br>"
Else
    Response.Write "d1 is NOT a valid date.<br>"
End If
%>