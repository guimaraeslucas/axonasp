

<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AxonASP - Error Handling Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .warning { background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 15px 0; border-radius: 4px; }
        .error-section { border-left: 4px solid #dc3545; padding: 15px; margin: 15px 0; background: #f8f9fa; border-radius: 4px; }
        hr { margin: 20px 0; border: none; border-top: 1px solid #ddd; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Error Handling Test</h1>
        <div class="intro">
            <p>This page demonstrates `Err.Raise` and `On Error Resume Next` behavior implemented by AxonASP.</p>
            <p>Two scenarios are shown: (1) using `On Error Resume Next` to capture errors, and (2) a hard raise that stops the script (enabled via <code>?stop=1</code>).</p>
        </div>

        <div class="error-section">
        <%
            Response.Write "<h3>1) On Error Resume Next (Err populated)</h3>"
            On Error Resume Next
            Err.Clear
            Err.Raise 100, "TestSource", "This is a test error (Resume Next)"
            Response.Write "Err.Number: " & Err.Number & "<br>"
            Response.Write "Err.Source: " & Err.Source & "<br>"
            Response.Write "Err.Description: " & Err.Description & "<br>"

            Response.Write "<hr>"

            Response.Write "<h3>2) Err.Raise without On Error Resume Next (use ?stop=1)</h3>"
            On Error GoTo 0
            If Request.QueryString("stop") = "1" Then
                Err.Clear
                Err.Raise 200, "TestSource2", "This will raise a runtime error and halt execution"
                Response.Write "This line will not execute if Err.Raise halts the script." 
            Else
                Response.Write "Pass ?stop=1 to the URL to trigger a hard raise that halts execution." 
            End If
        %>
        </div>
    </div>
</body>
</html>

