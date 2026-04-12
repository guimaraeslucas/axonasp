<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Error Handling Test</title>
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
        <h1>G3pix AxonASP - Error Handling Test</h1>
        <div class="intro">
            <p>Tests On Error Resume Next, error detection with Err object, Err.Clear, Err.Raise and error handling flow control.</p>
        </div>

<%
' Test Error Handling
Dim x, y, z

Response.Write "<div class='box'><h3>Testing On Error Resume Next</h3>"

On Error Resume Next

' Cause Division by Zero
x = 10
y = 0
z = x / y

If Err.Number <> 0 Then
    Response.Write "Caught Error: " & Err.Number & " - " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "No Error (Failed Test)<br>"
End If

Response.Write "Continuing execution...<br>"

' Another Error
Dim arr(2)
' Out of bounds
z = arr(5)

If Err.Number <> 0 Then
    Response.Write "Caught Array Error: " & Err.Number & " - " & Err.Description & "<br>"
    Err.Clear
End If

Response.Write "</div>"

Response.Write "<div class='box'><h3>Testing On Error GoTo 0</h3>"
On Error GoTo 0

' Should Crash (uncomment to test crash manually, but kept commented for safe run)
' z = x / y
Response.Write "Error handling disabled. (Skipping crash test)<br>"
Response.Write "</div>"

Response.Write "<div class='box'><h3>Testing Err.Raise</h3>"
On Error Resume Next
Err.Raise 999, "CustomScript", "Something went wrong"
Response.Write "Raised Error: " & Err.Number & " - " & Err.Source & " - " & Err.Description & "<br>"
Response.Write "</div>"

%>
    </div>
</body>
</html>
