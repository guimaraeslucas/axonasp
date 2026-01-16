<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Features & Operators Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .box { border-left: 4px solid #667eea; padding: 15px; margin-bottom: 15px; background: #f9f9f9; border-radius: 4px; }
        pre { background: #f4f4f4; padding: 10px; border-radius: 4px; overflow-x: auto; }
        code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; }
        form { margin: 15px 0; }
        input, button { padding: 8px 12px; margin: 5px 5px 5px 0; border: 1px solid #ddd; border-radius: 4px; }
        button { background: #667eea; color: #fff; cursor: pointer; }
        button:hover { background: #764ba2; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Features & Operators Test</h1>
        <div class="intro">
            <p>Tests mathematical operators, subroutines with parameters, Request object and form operations.</p>
        </div>

    <h2>1. Math Operators</h2>
    <%
        Dim a, b
        a = 10
        b = 5
        Response.Write "a=" & a & ", b=" & b & "<br>"
        Response.Write "a * b = " & (a * b) & "<br>"
        Response.Write "a - b = " & (a - b) & "<br>"
        Response.Write "100 / 2 * 5 = " & (100 / 2 * 5) & " (Should be 250)<br>"
        Response.Write "10 - 5 - 2 = " & (10 - 5 - 2) & " (Should be 3)<br>"
    %>

    <h2>2. Subroutines with Params</h2>
    <%
        Sub Calc(x, y)
            Response.Write "Inside Calc: x=" & x & ", y=" & y & "<br>"
            Response.Write "Product: " & (x * y) & "<br>"
        End Sub

        Response.Write "Calling Calc(6, 7)...<br>"
        Call Calc(6, 7)
    %>

    <h2>3. Request Object</h2>
    <p>Try appending ?name=Lucas to the URL.</p>
    <%
        Dim n
        n = Request.QueryString("name")
        If n = "" Then
            Response.Write "No name provided in QueryString.<br>"
        Else
            Response.Write "Hello, " & n & "! (from QueryString)<br>"
        End If
    %>

    <h3>Form Post</h3>
    <form method="POST">
        <input type="text" name="city" value="Sao Paulo">
        <input type="submit" value="Send">
    </form>
    <%
        Dim c
        c = Request.Form("city")
        If c <> "" Then
            Response.Write "You posted city: " & c & "<br>"
        End If
    %>
    </div>

    </div>
</body>
</html>

