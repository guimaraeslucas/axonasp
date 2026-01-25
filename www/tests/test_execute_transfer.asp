<%@ Language=VBScript %>
<%
Response.Write "<h3>Main Page Start</h3>"
Response.Write "<p>Calling Server.Execute('test_exec_target.asp')...</p>"
Server.Execute("test_exec_target.asp")
Response.Write "<p>Back in Main Page (Execute finished).</p>"

Response.Write "<hr>"

Response.Write "<p>Calling Server.Transfer('test_transfer_target.asp')...</p>"
Response.Write "<p>This text should NOT be seen if Transfer works (buffer cleared).</p>"
Server.Transfer("test_transfer_target.asp")
Response.Write "<p>FAIL: This text should NOT be seen (execution should stop).</p>"
%>
