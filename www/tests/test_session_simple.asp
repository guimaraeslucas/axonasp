<%
Session("test") = "Hello"
Response.Write "Session ID: " & Session.SessionID & "<br>"
Response.Write "Data: " & Session("test") & "<br>"
%>
