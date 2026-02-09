<%
' FastCGI Test Page
Response.ContentType = "text/html; charset=utf-8"
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AxonASP FastCGI Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        .success { color: green; }
        .info { background: #f0f0f0; padding: 10px; margin: 10px 0; border-radius: 4px; }
        h1 { color: #333; }
        h2 { color: #666; margin-top: 30px; }
        .test-result { margin: 10px 0; padding: 8px; border-left: 3px solid #4CAF50; background: #f9f9f9; }
    </style>
</head>
<body>
    <h1 class="success">✓ AxonASP FastCGI Mode Working!</h1>
    
    <h2>Server Information</h2>
    <div class="info">
        <strong>Server Name:</strong> <%= Request.ServerVariables("SERVER_NAME") %><br>
        <strong>Request Method:</strong> <%= Request.ServerVariables("REQUEST_METHOD") %><br>
        <strong>Script Name:</strong> <%= Request.ServerVariables("SCRIPT_NAME") %><br>
        <strong>Current Path:</strong> <%= Server.MapPath(".") %><br>
        <strong>Server Time:</strong> <%= Now() %><br>
    </div>

    <h2>Session Test</h2>
    <%
    If IsEmpty(Session("visit_count")) Then
        Session("visit_count") = 1
    Else
        Session("visit_count") = Session("visit_count") + 1
    End If
    %>
    <div class="test-result">
        <strong>Session ID:</strong> <%= Session.SessionID %><br>
        <strong>Visit Count:</strong> <%= Session("visit_count") %>
    </div>

    <h2>Application Test</h2>
    <%
    Application.Lock
    If IsEmpty(Application("total_visits")) Then
        Application("total_visits") = 1
    Else
        Application("total_visits") = Application("total_visits") + 1
    End If
    Application.Unlock
    %>
    <div class="test-result">
        <strong>Total Application Visits:</strong> <%= Application("total_visits") %>
    </div>

    <h2>Custom Library Test (G3JSON)</h2>
    <%
    Dim json
    Set json = Server.CreateObject("G3JSON")
    
    Dim testData
    Set testData = json.NewObject()
    testData("name") = "FastCGI Test"
    testData("status") = "working"
    testData("timestamp") = CStr(Now())
    
    Dim jsonString
    jsonString = json.Stringify(testData)
    %>
    <div class="test-result">
        <strong>JSON Output:</strong><br>
        <code><%= Server.HTMLEncode(jsonString) %></code>
    </div>

    <h2>Form Test</h2>
    <form method="post" action="">
        <input type="text" name="testInput" placeholder="Enter some text" value="<%= Request.Form("testInput") %>">
        <button type="submit">Submit</button>
    </form>
    <%
    If Request.Form("testInput") <> "" Then
        Response.Write "<div class='test-result'>"
        Response.Write "<strong>Form Data Received:</strong> " & Server.HTMLEncode(Request.Form("testInput"))
        Response.Write "</div>"
    End If
    %>

    <h2>Query String Test</h2>
    <%
    If Request.QueryString("test") <> "" Then
        Response.Write "<div class='test-result'>"
        Response.Write "<strong>Query String Received:</strong> test=" & Server.HTMLEncode(Request.QueryString("test"))
        Response.Write "</div>"
    Else
        Response.Write "<div class='info'>"
        Response.Write "<a href='?test=hello'>Click here to test query string</a>"
        Response.Write "</div>"
    End If
    %>

    <hr>
    <p><small>
        Powered by <strong>G3Pix AxonASP</strong> in FastCGI Mode<br>
        <a href="/tests/">View All Tests</a> | 
        <a href="/">Return to Home</a>
    </small></p>
</body>
</html>
