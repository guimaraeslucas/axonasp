<%
Response.Write "BEFORE_EXECUTE|"
Server.Execute "/tests/pieter/server_child.asp"
Response.Write "|AFTER_EXECUTE"
%>
