<%@ Language=VBScript %>
<%
Option Explicit

Function pre(value)
    Dim output
    output = Replace(value, vbTab, " ")
End Function

%>
<!DOCTYPE html>
<html>
<head>
    <title>Test Option Explicit</title>
</head>
<body>
<p><%= pre("test") %></p>
</body>
</html>
