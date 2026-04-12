<%
' Test file for 404 error handling
' This demonstrates both default and IIS modes
Response.ContentType = "text/html"
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>404 Error Handling Test - AxonASP</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background: #f5f5f5;
        }
        h1 {
            color: #333;
            border-bottom: 3px solid #667eea;
            padding-bottom: 10px;
        }
        .section {
            background: white;
            padding: 20px;
            margin: 20px 0;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .test-link {
            display: inline-block;
            background: #667eea;
            color: white;
            padding: 10px 20px;
            text-decoration: none;
            border-radius: 5px;
            margin: 5px;
            transition: background 0.3s;
        }
        .test-link:hover {
            background: #764ba2;
        }
        .info {
            background: #e8f4f8;
            padding: 15px;
            border-left: 4px solid #667eea;
            margin: 10px 0;
        }
        .code {
            background: #f4f4f4;
            padding: 10px;
            border-radius: 3px;
            font-family: 'Courier New', monospace;
            margin: 10px 0;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
        }
        th, td {
            padding: 10px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        th {
            background: #667eea;
            color: white;
        }
    </style>
</head>
<body>
    <h1>🔧 404 Error Handling Test Page</h1>
    
    <div class="section">
        <h2>Feature Overview</h2>
        <p>AxonASP now supports custom 404 error handling with two modes:</p>
        
        <table>
            <tr>
                <th>Mode</th>
                <th>Configuration</th>
                <th>Behavior</th>
            </tr>
            <tr>
                <td><strong>default</strong></td>
                <td><code>ERROR_404_MODE=default</code></td>
                <td>Serves static HTML from <code>errorpages/404.html</code></td>
            </tr>
            <tr>
                <td><strong>IIS</strong></td>
                <td><code>ERROR_404_MODE=IIS</code></td>
                <td>Reads <code>www/web.config</code> and executes custom ASP handler</td>
            </tr>
        </table>
    </div>
    
    <div class="section">
        <h2>Current Configuration</h2>
        <div class="info">
            <strong>Server Time:</strong> <%= Now() %><br>
            <strong>Server Software:</strong> AxonASP Server - G3pix Ltda<br>
            <strong>Current Page:</strong> <%= Request.ServerVariables("SCRIPT_NAME") %>
        </div>
    </div>
    
    <div class="section">
        <h2>Test Custom 404 Handlers</h2>
        <p>Click the links below to trigger 404 errors and see the custom handler in action:</p>
        
        <a href="/nonexistent-page.asp" class="test-link" target="_blank">
            Test 1: Missing ASP File
        </a>
        
        <a href="/fake/directory/file.html" class="test-link" target="_blank">
            Test 2: Deep Path Not Found
        </a>
        
        <a href="/random-<%= Int(Rnd() * 10000) %>.asp" class="test-link" target="_blank">
            Test 3: Random URL
        </a>
    </div>
    
    <div class="section">
        <h2>Configuration Instructions</h2>
        
        <h3>Step 1: Set ERROR_404_MODE in .env</h3>
        <div class="code">
# .env file<br>
ERROR_404_MODE=IIS  # or "default"
        </div>
        
        <h3>Step 2: Configure web.config (for IIS mode)</h3>
        <div class="code">
&lt;?xml version="1.0" encoding="UTF-8"?&gt;<br>
&lt;configuration&gt;<br>
&nbsp;&nbsp;&lt;system.webServer&gt;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&lt;httpErrors errorMode="Custom"&gt;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;error statusCode="404" path="/custom_404.asp" responseMode="ExecuteURL" /&gt;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&lt;/httpErrors&gt;<br>
&nbsp;&nbsp;&lt;/system.webServer&gt;<br>
&lt;/configuration&gt;
        </div>
        
        <h3>Step 3: Create Custom Error Handler (ASP)</h3>
        <div class="code">
&lt;%<br>
' custom_404.asp<br>
Response.ContentType = "text/html"<br>
%&gt;<br>
&lt;!DOCTYPE html&gt;<br>
&lt;html&gt;<br>
&nbsp;&nbsp;&lt;body&gt;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&lt;h1&gt;404 - Page Not Found&lt;/h1&gt;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&lt;p&gt;Requested: &lt;%= Request.ServerVariables("URL") %&gt;&lt;/p&gt;<br>
&nbsp;&nbsp;&lt;/body&gt;<br>
&lt;/html&gt;
        </div>
    </div>
    
    <div class="section">
        <h2>Implementation Details</h2>
        <ul>
            <li><strong>Web.config Parser:</strong> <code>server/webconfig_parser.go</code> handles XML parsing</li>
            <li><strong>Error Routing:</strong> <code>main.go</code> routes 404s based on configuration</li>
            <li><strong>ASP Execution:</strong> Custom error handlers execute with full ASP context</li>
            <li><strong>Security:</strong> Paths are validated against root directory</li>
            <li><strong>Fallback:</strong> Automatically falls back to default on errors</li>
        </ul>
    </div>
    
    <div class="section">
        <h2>✅ Test Results</h2>
        <p>If you're seeing this page, the server is working correctly!</p>
        <div class="info">
            Test the 404 handlers using the links above to verify custom error handling is functioning.
        </div>
    </div>
    
    <hr>
    <p style="text-align: center; color: #999; font-size: 12px;">
        AxonASP Server &copy; 2026 G3pix Ltda - Licensed under MPL 2.0
    </p>
</body>
</html>
