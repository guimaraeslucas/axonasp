<%
' Custom 404 error handler for AxonASP
Response.ContentType = "text/html"
%>
<!DOCTYPE html>
<html lang="en">
    <!--
        AxonASP Server
        Copyright (C) 2026 G3pix Ltda. All rights reserved.
    -->

    <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>404 - Not Found - AxonASP Server</title>
        <link rel="icon" href="/favicon.ico" />
        <link rel="shortcut icon" href="/favicon.ico" />
        <link rel="stylesheet" href="/css/axonasp.css" />
        <%
        Dim ax
        Set ax = Server.CreateObject("G3Axon.Functions")
        %>
    </head>

    <body class="error-page">
        <div id="header">
            <div class="logo">
                <img src="<%= ax.AxGetLogo() %>" alt="AxonASP" width="43" />
            </div>
            <h1>AxonASP Server</h1>
        </div>

        <div class="error-shell">
            <div class="error-card error-card-wide">
                <h1>404 - Not Found</h1>
                <p>
                    The page you requested cannot be found. It may have been removed,
                    renamed, or is temporarily unavailable.
                </p>

                <h2>Technical Information</h2>
                <table class="error-table">
                    <tbody>
                        <tr>
                            <th>Requested URL</th>
                            <td><%= Request.ServerVariables("URL") %></td>
                        </tr>
                        <tr>
                            <th>Server Time</th>
                            <td><%= Now() %></td>
                        </tr>
                        <tr>
                            <th>Error Handler</th>
                            <td>Custom ASP Handler (web.config)</td>
                        </tr>
                    </tbody>
                </table>

                <p>
                    Check the URL for typing errors or return to the
                    <a href="/">home page</a>.
                </p>

                <div class="error-footer">
                    &copy; 2026 G3Pix AxonASP. All rights reserved.<br />
                    For support, visit
                    <a href="https://g3pix.com.br/axonasp">https://g3pix.com.br/axonasp</a>
                </div>
            </div>
        </div>
    </body>

</html>