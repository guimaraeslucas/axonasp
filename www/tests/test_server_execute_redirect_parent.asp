<%@ Language="VBScript" %>
<!--
Test: Server.Execute Response.Redirect propagation
Verifies that Response.Redirect inside a child file executed via Server.Execute
propagates to the parent and terminates the entire page.

Expected behavior:
  - "BEFORE|" may appear in output (written before Server.Execute)
  - Response.Redirect inside the child fires
  - "|AFTER" must NOT appear in output
  - HTTP status must be 302 Found
-->
<%
Response.Write "BEFORE|"
Server.Execute "test_server_execute_redirect_child.asp"
Response.Write "|AFTER"
%>