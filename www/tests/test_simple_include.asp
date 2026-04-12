<%
@ Language = VBScript
%>
<!DOCTYPE html>
<html>
    <head>
        <title>Simple AxInclude Test</title>
    </head>
    <body>
        <h1>Simple AxInclude Test</h1>
        <p>Before Include</p>
        <%
        Response.Write "About to call AxInclude..." & "<br>"
        AxInclude("config.inc")
        Response.Write "After AxInclude call" & "<br>"
        %>
        <p>After Include</p>
    </body>
</html>
