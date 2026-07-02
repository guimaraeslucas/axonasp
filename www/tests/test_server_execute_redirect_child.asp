<%@ Language="VBScript" %>
<%
' Child file for test_server_execute_redirect_parent.asp
' This file is executed via Server.Execute and calls Response.Redirect.
' The redirect must propagate to the parent and terminate the entire page.
Sub Inner(target, msg)
    Response.Redirect "target.asp?act=" & target & "&msg=" & Server.URLEncode(msg)
End Sub

' Bare call with arguments - this must terminate the entire page
Inner "index", "success"
%>