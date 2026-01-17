<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Date Literals Test</title>
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
        <h1>G3pix AxonASP - Date Literals Test</h1>
        <div class="intro">
            <p>Tests date literal syntax with # delimiters and date arithmetic operations.</p>
        </div>
        <%
Response.Write "<h3>Testing Date Literals</h3>"
Dim d
d = #1/1/2023#
Response.Write "Date Literal #1/1/2023# type: " & TypeName(d) & "<br>"
Response.Write "Date Literal #1/1/2023# value: " & d & "<br>"

Dim d2
d2 = #2023-12-31#
Response.Write "Date Literal #2023-12-31# value: " & d2 & "<br>"

Dim diff
diff = DateDiff("d", d, d2)
Response.Write "Days between: " & diff & "<br>"

Response.Write "<h3>Testing Date in Math</h3>"
' Note: In this interpreter, + operator defaults to string concatenation for non-integers
Dim d3
d3 = d + " is a date"
Response.Write "Concatenation: " & d3 & "<br>"
%>
    </div>
</body>
</html>