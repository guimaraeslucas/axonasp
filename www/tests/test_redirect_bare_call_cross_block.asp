<%@ Language="VBScript" %>

<!-- #include file="test_redirect_bare_call_helper.asp" -->
<%
Response.Write "before<br>"
' Bare call WITH arguments crossing a block boundary - MUST terminate script
RedirectMsg "index", "test"
Response.Write "after (should NOT appear)<br>"
%>