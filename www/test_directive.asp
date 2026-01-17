<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3 AxonASP - Directive Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .result { background: #f9f9f9; border-left: 4px solid #667eea; padding: 15px; margin: 15px 0; border-radius: 4px; }
        .success { background: #e8f5e9; border-left: 4px solid #4caf50; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3 AxonASP - Directive Test</h1>
        <div class="intro">
            <p>Tests ASP directives like <code>&lt;%@ Language=VBScript %&gt;</code></p>
        </div>

        <h2>Test Results</h2>
        <div class="result success">
            <%
                Response.Write "<p><strong>Success!</strong> This page uses the <code>&lt;%@ Language=VBScript %&gt;</code> directive.</p>"
                Response.Write "<p>Current Time: " & Now() & "</p>"
                Response.Write "<p>VBScript is working correctly!</p>"
            %>
        </div>

        <h2>Variable Test</h2>
        <div class="result">
            <%
                Dim testVar
                testVar = "Directive support is working!"
                Response.Write "<p>" & testVar & "</p>"
            %>
        </div>

        <h2>Loop Test</h2>
        <div class="result">
            <%
                Dim i
                Response.Write "<p>Numbers: "
                For i = 1 To 5
                    Response.Write i & " "
                Next
                Response.Write "</p>"
            %>
        </div>
    </div>
</body>
</html>
