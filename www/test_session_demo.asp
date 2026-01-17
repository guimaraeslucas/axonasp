<!DOCTYPE html>
<html>
<head>
    <title>Session File Storage Demo</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .container { max-width: 900px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; }
        h1 { color: #333; border-bottom: 3px solid #667eea; padding-bottom: 10px; }
        .demo-box { background: #f9f9f9; border: 1px solid #ddd; padding: 20px; margin: 20px 0; border-radius: 4px; }
        .code { background: #f0f0f0; padding: 15px; border-radius: 4px; font-family: monospace; overflow-x: auto; margin: 10px 0; }
        .info { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin: 15px 0; border-radius: 4px; }
        .success { background: #e8f5e9; border-left: 4px solid #4caf50; padding: 15px; margin: 15px 0; border-radius: 4px; color: #2e7d32; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background: #667eea; color: white; }
        tr:hover { background: #f5f5f5; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Session File Storage Implementation</h1>
        
        <div class="info">
            <strong>ℹ️ How it works:</strong> Each ASP session is stored in a separate JSON file in the <code>temp/session/</code> directory. 
            Data persists across server restarts and is automatically cleaned up after 20 minutes of inactivity.
        </div>
        
        <div class="demo-box">
            <h2>1. Session Variables Created</h2>
            <p>The following variables are stored in the current session:</p>
            <table>
                <tr>
                    <th>Variable Name</th>
                    <th>Value</th>
                    <th>Type</th>
                </tr>
                <%
                    ' Set various session variables
                    Session("page_name") = "Session Demo"
                    Session("user_name") = "Guest User"
                    Session("visit_time") = Now()
                    Session("visit_count") = 1
                    
                    ' Display them in a table
                    Response.Write "<tr><td>page_name</td><td>" & Session("page_name") & "</td><td>String</td></tr>"
                    Response.Write "<tr><td>user_name</td><td>" & Session("user_name") & "</td><td>String</td></tr>"
                    Response.Write "<tr><td>visit_time</td><td>" & Session("visit_time") & "</td><td>DateTime</td></tr>"
                    Response.Write "<tr><td>visit_count</td><td>" & Session("visit_count") & "</td><td>Integer</td></tr>"
                %>
            </table>
        </div>
        
        <div class="demo-box">
            <h2>2. Session ID & File Location</h2>
            <table>
                <tr>
                    <th>Property</th>
                    <th>Value</th>
                </tr>
                <%
                    Dim sessionID
                    sessionID = Session.SessionID
                    Response.Write "<tr><td>Session.SessionID</td><td>" & sessionID & "</td></tr>"
                    Response.Write "<tr><td>File Path</td><td>temp/session/" & sessionID & ".json</td></tr>"
                %>
            </table>
        </div>
        
        <div class="success">
            <strong>✓ Session File Created:</strong> The session data for this request has been saved to 
            <code>temp/session/<% Response.Write Session.SessionID %>.json</code>
        </div>
        
        <div class="demo-box">
            <h2>3. Session File Format</h2>
            <p>The session file is stored in JSON format for easy debugging and inspection:</p>
            <div class="code">{
  "id": "<% Response.Write Session.SessionID %>",
  "data": {
    "page_name": "<% Response.Write Session("page_name") %>",
    "user_name": "<% Response.Write Session("user_name") %>",
    "visit_time": "<% Response.Write Session("visit_time") %>",
    "visit_count": <% Response.Write Session("visit_count") %>
  },
  "created_at": "2026-01-17T...",
  "last_accessed": "2026-01-17T...",
  "timeout": 20
}</div>
        </div>
        
        <div class="demo-box">
            <h2>4. Key Features</h2>
            <ul>
                <li><strong>File-based persistence:</strong> Session data survives server restarts</li>
                <li><strong>Automatic cleanup:</strong> Expired sessions (> 20 mins) are automatically removed</li>
                <li><strong>HTTP Cookie:</strong> Session ID is sent via ASPSESSIONID cookie</li>
                <li><strong>JSON storage:</strong> Human-readable format for debugging</li>
                <li><strong>Case-insensitive:</strong> Session keys are case-insensitive as per ASP spec</li>
                <li><strong>Full compatibility:</strong> Works with ASP syntax: Session("key") and Session.SessionID</li>
            </ul>
        </div>
        
        <div class="info">
            <strong>📁 Directory Structure:</strong>
            <br><code>go-asp/</code>
            <br>└── <code>temp/</code>
            <br>&nbsp;&nbsp;&nbsp;&nbsp;└── <code>session/</code>
            <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;├── <code>ASP1768693765685678300.json</code>
            <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;├── <code>ASP1768693772678633000.json</code>
            <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── <code>... (other session files)</code>
        </div>
    </div>
</body>
</html>
