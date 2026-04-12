<%@ Language=VBScript %>
<%
Option Explicit

Call TestOnError()

Sub TestOnError()
    On Error Resume Next
    
    Dim x
    x = "test"
%>
<html>
<body>
<h1>Test On Error</h1>
<p>Value: <%=x%></p>
</body>
</html>
<%
End Sub
%>
