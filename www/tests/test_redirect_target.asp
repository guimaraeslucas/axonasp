<%@ Language=VBScript %>
<%
Response.Write "Redirecting..."
Response.Redirect "test_fixes_v2.asp"
Response.Write "This text should NOT appear if Redirect works correctly (execution stopped)."
%>
