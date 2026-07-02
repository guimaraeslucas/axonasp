<%@ Language="VBScript" %>
<%
' Child file for test_server_execute_end_parent.asp
' This file is executed via Server.Execute and calls Response.End.
' Response.End must propagate to the parent and terminate the entire page.
Response.End
Response.Write "this should NOT appear"
%>