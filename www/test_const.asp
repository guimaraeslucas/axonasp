<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Constants Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .box { border-left: 4px solid #667eea; padding: 15px; margin-bottom: 15px; background: #f9f9f9; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Constants Test</h1>
        <div class="intro">
            <p>Tests constant declaration with Const keyword and immutability enforcement.</p>
        </div>
        <div class="box">

<%
Response.Write "<h3>Testing Constants</h3>"

Const MyNum = 100
Const MyStr = "AxonASP"

Response.Write "MyNum: " & MyNum & "<br>"
Response.Write "MyStr: " & MyStr & "<br>"

Dim x
x = MyNum + 50
Response.Write "Calculation with constant (100 + 50): " & x & "<br>"

Response.Write "<h4>Testing Illegal Assignment</h4>"
Response.Write "Attempting to assign to 'MyNum'...<br>"

On Error Resume Next
MyNum = 200
If Err.Number <> 0 Then
    Response.Write "Error Caught: " & Err.Description & " (Code: " & Err.Number & ")<br>"
    Err.Clear
Else
    Response.Write "No Error! (This is bad, reassignment should fail)<br>"
End If
On Error Goto 0

Response.Write "Value after attempt: " & MyNum & "<br>"
%>
        </div>
    </div>
</body>
</html>