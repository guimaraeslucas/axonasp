<%
    @ Language = VBScript
%>
<%
Option Explicit

Function pre(value)
    Dim Output
    Output = Replace(value, vbTab, " ")
End Function

%>
<!DOCTYPE html>
<html>
    <head>
        <title>Test Option Explicit</title>
    </head>
    <body>
        <p>
            <%= pre("test") %>
            - Must return Empty
        </p>
    </body>
</html>
