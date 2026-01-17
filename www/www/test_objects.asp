<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Core Objects Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .box { border-left: 4px solid #667eea; padding: 15px; margin: 15px 0; background: #f9f9f9; border-radius: 4px; }
        form { margin: 15px 0; }
        input, button { padding: 8px 12px; margin: 5px 5px 5px 0; border: 1px solid #ddd; border-radius: 4px; }
        button { background: #667eea; color: #fff; cursor: pointer; }
        button:hover { background: #764ba2; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Core Objects Test</h1>
        <div class="intro">
            <p>Tests Request, Response, Server and Session objects.</p>
        </div>

<h2>Request Object</h2>
Server Variable (HTTP_HOST): <%= Request.ServerVariables("HTTP_HOST") %><br>
QueryString (msg): <%= Request.QueryString("msg") %><br>
Cookie (test_cookie): <%= Request.Cookies("test_cookie") %><br>

<h2>Response Object</h2>
<% 
    Response.Cookies("test_cookie") = "CookieValue"
    Response.ContentType = "text/html"
    Response.Status = "200 OK"
    ' Response.Buffer = True ' Not observable directly here
%>
Cookies Set. Refresh to see Request.Cookies value above.<br>

<h2>Server Object</h2>
HTMLEncode("&<>"): <%= Server.HTMLEncode("&<>") %><br>
URLEncode("Hello World"): <%= Server.URLEncode("Hello World") %><br>
MapPath("test.asp"): <%= Server.MapPath("test.asp") %><br>

<h2>Session Object</h2>
<% 
    Session.Timeout = 30
    Session("User") = "Admin"
%>
Session ID: <%= Session.SessionID %><br>
Session Timeout: <%= Session.Timeout %> (Set to 30, but reading back might need property impl)<br>
Session("User"): <%= Session("User") %><br>

<h2>Redirect Test</h2>
<a href="test_redirect.asp">Clique aqui para testar Response.Redirect</a>
    </div>
</body>
</html>

