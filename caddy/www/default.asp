<% @Language="VBScript" %>
<%
Application.Lock
Application("RequestsCount") = Application("RequestsCount") + 1
currentCount = Application("RequestsCount")
Application.Unlock
%>
<!DOCTYPE html>
<html>

    <head>
        <title>AxonASP Caddy Test</title>
        <style>
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background: radial-gradient(circle at top left, #1e1e2e, #11111b);
                color: #cdd6f4;
                margin: 0;
                padding: 0;
                display: flex;
                justify-content: center;
                align-items: center;
                min-height: 100vh;
            }

            .card {
                background: rgba(49, 50, 68, 0.4);
                backdrop-filter: blur(12px);
                -webkit-backdrop-filter: blur(12px);
                border: 1px solid rgba(255, 255, 255, 0.08);
                padding: 35px 40px;
                border-radius: 16px;
                box-shadow: 0 20px 40px rgba(0, 0, 0, 0.4);
                max-width: 600px;
                width: 90%;
            }

            h1 {
                color: #f5c2e7;
                font-size: 2rem;
                margin-top: 0;
                margin-bottom: 25px;
                border-bottom: 1px solid rgba(255, 255, 255, 0.1);
                padding-bottom: 15px;
                text-align: center;
                letter-spacing: 1px;
            }

            .info {
                margin: 18px 0;
                font-size: 1.15rem;
                display: flex;
                justify-content: space-between;
                border-bottom: 1px dashed rgba(255, 255, 255, 0.05);
                padding-bottom: 8px;
            }

            .label {
                color: #a6adc8;
            }

            .value {
                font-weight: 600;
                color: #89b4fa;
            }

            .footer {
                margin-top: 30px;
                text-align: center;
                font-size: 0.85rem;
                color: #6c7086;
            }
        </style>
    </head>

    <body>
        <div class="card">
            <h1>❖ AxonASP inside Caddy</h1>
            <div class="info">
                <span class="label">Application Name:</span>
                <span class="value"><%= Application("SiteName") %></span>
            </div>
            <div class="info">
                <span class="label">Session Started:</span>
                <span class="value"><%= Session("UserStarted") %></span>
            </div>
            <div class="info">
                <span class="label">Current Server Time:</span>
                <span class="value"><%= Now() %></span>
            </div>
            <div class="info">
                <span class="label">Request URL:</span>
                <span class="value"><%= Request.ServerVariables("URL") %></span>
            </div>
            <div class="info">
                <span class="label">Page Views Counter:</span>
                <span class="value"><%= currentCount %></span>
            </div>
            <div class="info">
                <span class="label">User Agent:</span>
                <span class="value"><%= Request.ServerVariables("HTTP_USER_AGENT") %></span>
            </div>
            <div class="footer">
                Powered by AxonASP & Caddy Server
            </div>
        </div>
    </body>

</html>