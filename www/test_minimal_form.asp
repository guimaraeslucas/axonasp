<%
If Request.Form("submitted") = "1" Then
    Response.Write("YES")
Else
    Response.Write("NO")
End If
%>
