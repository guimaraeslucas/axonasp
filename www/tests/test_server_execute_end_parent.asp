<%@ Language="VBScript" %>
<!--
Test: Server.Execute Response.End propagation
Verifies that Response.End inside a child file executed via Server.Execute
propagates to the parent and terminates the entire page.

Expected behavior:
  - "BEFORE|" must appear in output (written before Server.Execute, flushed by Response.End)
  - "|AFTER" must NOT appear in output
-->
<%
Response.Write "BEFORE|"
Server.Execute "test_server_execute_end_child.asp"
Response.Write "|AFTER"
%>