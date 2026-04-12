<%
@ Language = "VBScript" CodePage = "65001" EnableSessionState = "False"
%>
<!DOCTYPE html>
<html>
    <head>
        <title>ASP Directives Test</title>
    </head>
    <body>
        <h1>ASP Directives Test</h1>
        <p>
            Response.CodePage:
            <%= Response.CodePage %>
        </p>
        <p>Session disabled directive applied before page execution.</p>
    </body>
</html>
