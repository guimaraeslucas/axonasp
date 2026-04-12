<%@ Language=VBScript %>
<%
Call TestSub()

Sub TestSub()
%>
<html>
<body>
<h1>Test Sub with HTML</h1>
<% If True Then %>
  <p>This is inside Sub and If</p>
<% End If %>
</body>
</html>
<%
End Sub
%>
