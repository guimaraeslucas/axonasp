<%@ Language=VBScript %>
<%
Response.Write "<div style='background:#eef; padding:10px; border:1px solid blue;'>"
Response.Write "<strong>Inside Executed Page</strong><br>"
Response.Write "ScriptTimeout: " & Server.ScriptTimeout & "<br>"
Response.Write "</div>"
%>
