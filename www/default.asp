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
        <title>G3Pix AxonASP Server - Active</title>
        <link rel="stylesheet" href="css/axonasp.css" />
    </head>
    <body>
        <div id="header">
            <div class="logo">
                <%
                Dim ax
                Set ax = Server.CreateObject("G3Axon.Functions")
                %>
                <img
                    src="<%= ax.AxGetLogo() %>"
                    alt="AxonASP"
                    width="43"
                />
            </div>
            <h1>AxonASP Server</h1>
        </div>

        <div id="main-container">
            <div id="sidebar">
                <div class="section-title">Navigation</div>
                <ul>
                    <li><a href="default.asp">Home Page</a></li>
                    <li><a href="/manual">Software Manual</a></li>
                    <li>
                        <a href="/database-convert">Database Conversion Tool</a>
                    </li>
                    <li>
                        <a href="/project-builder">ASP Project Builder</a>
                    </li>
                    <li>
                        <a href="https://g3pix.com.br/axonasp" target="_blank"
                            >AxonASP web-site</a
                        >
                    </li>
                </ul>

                <div class="section-title" style="margin-top: 20px">
                    Resources
                </div>
                <ul>
                    <li><a href="tests/default.asp">System Tests</a></li>
                    <li>
                        <a href="rest/default.asp">REST example</a>
                    </li>
                    <li>
                        <a href="restful/default.asp">RESTful example</a>
                    </li>
                    <li>
                        <a href="mvc/default.asp">MVC example</a>
                    </li>
                    <li>
                        <a href="mvvm/default.asp">MVVM example</a>
                    </li>
                    <li>
                        <a href="jscript-api/default.asp">JScript API example</a>
                    </li>
                </ul>
            </div>

            <div id="content">
                <h1>It Works!</h1>
                <p>
                    Your AxonASP Server is successfully running and processing
                    Classic ASP / VBScript requests.
                </p>

                <div
                    style="
                        background-color: #ffffe1;
                        border: 1px solid #aca899;
                        padding: 10px;
                        margin: 15px 0;
                        font-size: 12px;
                    "
                >
                    <span class="status-badge"></span>
                    <strong>System Status:</strong> The server engine is active
                    and stable.
                </div>

                <p>
                    AxonASP provides a robust environment for hosting legacy
                    applications with modern performance. Below is the current
                    server configuration for your reference.
                </p>

                <h2>Server Environment Status</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Property</th>
                            <th>Value</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td><strong>Server Product</strong></td>
                            <td>G3Pix AxonASP Server</td>
                        </tr>
                        <tr>
                            <td><strong>Engine Version</strong></td>
                            <td>
                                <%= ax.AxEngineName() %>
                                /
                                <%= ax.AxVersion() %>
                            </td>
                        </tr>
                        <tr>
                            <td><strong>Current Date</strong></td>
                            <td><%= FormatDateTime(Now, vbLongDate) %></td>
                        </tr>
                        <tr>
                            <td><strong>Current Time</strong></td>
                            <td><%= FormatDateTime(Now, vbLongTime) %></td>
                        </tr>
                        <tr>
                            <td><strong>Currency Sample</strong></td>
                            <td><%= FormatCurrency(12345.67, 2) %></td>
                        </tr>
                        <tr>
                            <td><strong>Configured Locale (LCID)</strong></td>
                            <td><%= ax.AxGetConfig("global.default_mslcid") %></td>
                        </tr>
                        <tr>
                            <td><strong>Configured Timezone</strong></td>
                            <td><%= ax.AxGetConfig("global.default_timezone") %></td>
                        </tr>
                        <tr>
                            <td><strong>Active Session ID</strong></td>
                            <td><%= Session.SessionID %></td>
                        </tr>
                    </tbody>
                </table>

                <h2>Next Steps</h2>
                <p>
                    Your server is live and ready for real workloads. Pick your
                    next mission and start building now:
                </p>
                <ul>
                    <li>
                        Open the <a href="/manual">Software Manual</a> and try
                        one example from our G3 libraries.
                    </li>
                    <li>
                        Launch the
                        <a href="/asp-code-generator">ASP Code Generator</a> and
                        scaffold your next Classic ASP module in minutes.
                    </li>
                    <li>
                        Or convert your existing Access Database to SQLite using
                        our <a href="/database-convert/">migration tool</a>.
                    </li>
                </ul>

                <div style="margin-top: 20px; text-align: right">
                    <img src="PoweredBy.png" alt="Powered by AxonASP" />
                </div>
            </div>
            <div id="status-bar">
                <%= Now() %>
            </div>
        </div>
    </body>
</html>
