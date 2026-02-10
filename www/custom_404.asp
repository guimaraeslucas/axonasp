<%
' Test file for custom 404 error handling
' This file is used by web.config as the custom 404 handler
Response.ContentType = "text/html"
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>404 - Page Not Found</title>
	<link rel="icon" href="favicon.ico">
	<link rel="shortcut icon" href="favicon.ico">
    <style>
        body {
            font-family: Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            margin: 0;
            padding: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
        }
        .error-container {
            background: white;
            border-radius: 10px;
            padding: 40px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            text-align: center;
            max-width: 600px;
        }
        h1 {
            color: #667eea;
            font-size: 72px;
            margin: 0;
            font-weight: bold;
        }
        h2 {
            color: #333;
            font-size: 24px;
            margin: 20px 0;
        }
        p {
            color: #666;
            font-size: 16px;
            line-height: 1.6;
        }
        .btn {
            display: inline-block;
            background: #667eea;
            color: white;
            padding: 12px 30px;
            text-decoration: none;
            border-radius: 5px;
            margin-top: 20px;
            transition: background 0.3s;
        }
        .btn:hover {
            background: #764ba2;
        }
        .info {
            margin-top: 30px;
            padding: 15px;
            background: #f0f0f0;
            border-radius: 5px;
            font-size: 14px;
        }
        .server-info {
            color: #999;
            font-size: 12px;
            margin-top: 20px;
        }
    </style>
</head>
<body>
    <div class="error-container">
        <h1>404</h1>
        <h2>Page Not Found</h2>
        <p>The page you are looking for might have been removed, had its name changed, or is temporarily unavailable.</p>
        
        <div class="info">
            <strong>Requested URL:</strong> <%= Request.ServerVariables("URL") %><br>
            <strong>Server Time:</strong> <%= Now() %><br>
            <strong>Error Handler:</strong> Custom ASP (web.config)
        </div>
        
        <a href="/" class="btn">Go to Homepage</a>
        
        <div class="server-info">
            Powered by AxonASP Server - G3pix Ltda
        </div>
    </div>
</body>
</html>
