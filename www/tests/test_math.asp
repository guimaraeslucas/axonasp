<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Math Operators Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Math Operators Test</h1>
        <div class="intro">
            <p>Tests exponentiation operator (^) with various numeric values.</p>
        </div>
        <%
Response.Write "<h3>Math Operator Test</h3>"
Dim a, b, c
a = 2
b = 3
c = a ^ b
Response.Write "2 ^ 3 = " & c & "<br>"

Dim d
d = 10 ^ 2
Response.Write "10 ^ 2 = " & d & "<br>"

Dim e
e = 2 ^ 0.5
Response.Write "2 ^ 0.5 = " & e & "<br>"
%>
    </div>
</body>
</html>