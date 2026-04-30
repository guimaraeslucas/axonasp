<%
' Custom 404 error handler for AxonASP
Response.ContentType = "text/html"
%>
<!DOCTYPE html>
<html lang="en">
    <!--
        
        AxonASP Server
        Copyright (C) 2026 G3pix Ltda. All rights reserved.
        
        Developed by Lucas Guimarães - G3pix Ltda
        Contact: https://g3pix.com.br/
        Project URL: https://g3pix.com.br/axonasp
        
        This Source Code Form is subject to the terms of the Mozilla Public
        License, v. 2.0. If a copy of the MPL was not distributed with this
        file, You can obtain one at https://mozilla.org/MPL/2.0/.
        
        Attribution Notice:
        If this software is used in other projects, the name "AxonASP Server"
        must be cited in the documentation or "About" section.
        
        Contribution Policy:
        Modifications to the core source code of AxonASP Server must be
        made available under this same license terms.
        
        -->
    <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>404 - Not Found - AxonASP Server</title>
        <link rel="icon" href="/favicon.ico" />
        <link rel="shortcut icon" href="/favicon.ico" />
        <%
        Dim ax
        Set ax = Server.CreateObject("G3Axon.Functions")

        %>
        <style>
            <%= ax.axgetdefaultcss() %>
        </style>
    </head>
    <body>
        <div id="header">
            <div class="logo">
                <img
                    src="<%= ax.AxGetLogo() %>"
                    alt="AxonASP"
                    width="43"
                />
            </div>
            <h1>AxonASP Server</h1>
        </div>
        <div id="main-container">
            <div id="content">
                <h1>404 - Not Found</h1>
                <p>
                    The page you are looking for cannot be found. It might have
                    been removed, had its name changed, or is temporarily
                    unavailable.
                </p>

                <h2>Technical Information (for support personnel)</h2>
                <table style="width: 0">
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
                    Please check the URL for typos or click the button below to
                    return to the home page.
                </p>
                <a href="/" class="btn-link">Go to Homepage</a>

                <div class="footer">
                    &copy; 2026 G3pix Ltda. AxonASP Server. All rights
                    reserved.<br />
                    For technical support, visit
                    <a href="https://g3pix.com.br/axonasp"
                        >https://g3pix.com.br/axonasp</a
                    >
                </div>
            </div>
        </div>
    </body>
</html>
