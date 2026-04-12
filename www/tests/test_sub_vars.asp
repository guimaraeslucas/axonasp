<%@ Language=VBScript %>
<%
Option Explicit

Dim mode
mode = "html"

If mode = "html" Then
    Call RenderPage()
End If

Sub RenderPage()
    Dim x, y
    x = "value1"
    y = "value2"
%>
<html>
<body>
<h1>Test Variables</h1>
<p>X = <%=x%></p>
<p>Y = <%=y%></p>
<% If True Then %>
  <p>Inside If</p>
<% End If %>
</body>
</html>
<%
End Sub
%>
