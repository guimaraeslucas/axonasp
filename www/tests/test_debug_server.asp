<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Server Object Debug Test</title>
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
        <h1>G3pix AxonASP - Server Object Debug Test</h1>
        <div class="intro">
            <p>Debug and inspection of Server object functionality with CreateObject and File operations.</p>
        </div>
        <%
Response.Write "<h3>Debug Server Object</h3>"

Dim fs
Response.Write "1. Calling CreateObject...<br>"
Set fs = Server.CreateObject("G3FILES")

Response.Write "2. Checking IsObject(fs)...<br>"
If IsObject(fs) Then
    Response.Write "fs IS an Object.<br>"
Else
    Response.Write "fs is NOT an Object. (Type: " & TypeName(fs) & ")<br>"
End If

Response.Write "3. Calling fs.Write...<br>"
' Use a file path that is definitely safe
Dim ret
ret = fs.Write("debug.txt", "test")

Response.Write "4. Write Result: " & ret & "<br>"

Dim content
content = fs.Read("debug.txt")
Response.Write "5. Read Content: " & content & "<br>"
%>
    </div>
</body>
</html>
