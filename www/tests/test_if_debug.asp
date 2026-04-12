<%
@ Language = VBScript
%>
<html>
    <body>
        <h1>Test If Block</h1>
        <%
        If True Then
        %>
        <p>This is TRUE</p>
        <%
        Else
        %>
        <p>This is FALSE</p>
        <%
        End If
        %>
        <p>Outside If</p>
    </body>
</html>
