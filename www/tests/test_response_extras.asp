<%
Response.Write "<h3>Response Extras Test</h3>"

' Test PICS
Response.PICS = "(PICS-1.1 ""http://www.rsac.org/ratingsv01.html"" l gen true comment ""RSACi North America Server"" by ""inet@microsoft.com"" for ""http://www.microsoft.com"" on ""1997.06.30T14:21-0500"" r (n 0 s 0 v 0 l 0))"
Response.Write "PICS Property Set.<br>"
Response.Write "PICS Value: " & Response.PICS & "<br>"

' Test CacheControl / Pragma
Response.Expires = 0
Response.CacheControl = "Private"
Response.Write "CacheControl set to Private, Expires 0. Check Headers for Pragma: no-cache.<br>"

' Test IsClientConnected
If Response.IsClientConnected Then
    Response.Write "Client is connected.<br>"
Else
    Response.Write "Client is NOT connected.<br>"
End If
%>