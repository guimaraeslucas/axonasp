<%
Response.Write "Start test<br>"

On Error GoTo ErrorHandler

Response.Write "Before error (should print)<br>"
x = 1 / 0
Response.Write "After error (should NOT print)<br>"

' Explicit label syntax with colon
GoTo SkipHandler

ErrorHandler:
Response.Write "In error handler (should print)<br>"
Response.Write "Error: " & Err.Number & "<br>"

SkipHandler:
Response.Write "After handler (should print)<br>"

Response.Write "End test<br>"
%>
