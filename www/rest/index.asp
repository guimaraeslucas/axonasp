<%
' Test Web Config Rewrite Conditions
' This test should demonstrate that rest/something is rewritten to rest/index.asp?route=something
' IF the file doesn't exist.

Dim route
route = Request.QueryString("route")

If route <> "" Then
    Response.Write "SUCCESS: Route is " & route
Else
    Response.Write "FAILURE: No route found"
End If
%>
