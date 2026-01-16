<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Redirect Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .info { background: #d1ecf1; border-left: 4px solid #17a2b8; padding: 15px; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="info">
            <p>You will be redirected to the default page...</p>
        </div>
    </div>
<%
    Response.Redirect("lucas.asp")
%>
</body>
</html>

