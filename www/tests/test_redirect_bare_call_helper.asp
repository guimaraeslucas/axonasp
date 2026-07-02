<%@ Language="VBScript" %>
<%
' Helper file for test_redirect_bare_call_cross_block.asp
' Defines a Sub that does Response.Redirect when called.
Sub RedirectMsg(act, msg)
    Response.Redirect "target.asp?act=" & act & "&msg=" & Server.URLEncode(msg)
End Sub
%>