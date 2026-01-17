<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - File Operations Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .code-block { background: #f4f4f4; padding: 10px; margin: 10px 0; border-radius: 4px; word-break: break-all; }
        code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; }
        .success { color: #28a745; }
        .error { color: #dc3545; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - File Operations Test</h1>
        <div class="intro">
            <p>Tests file system functions: read, write, delete and list.</p>
        </div>

    <%
        Dim fs
        Set fs = Server.CreateObject("G3FILES")

        Dim path, content
        path = "demo_file.txt"
        
        ' 1. Write Text
        Response.Write "<h3>1. Writing File...</h3>"
        If fs.Write(path, "Hello World from GoLang API!" & vbCrLf) Then
            Response.Write "Success writing to " & path & "<br>"
        Else
            Response.Write "Error writing file.<br>"
        End If

        ' 2. Append Text
        Response.Write "<h3>2. Appending Text...</h3>"
        fs.Append(path, "This is a second line appended." & vbCrLf)
        Response.Write "Appended.<br>"

        ' 3. Check Existence
        Response.Write "<h3>3. Checking Existence...</h3>"
        If fs.Exists(path) Then
            Response.Write "File exists!<br>"
        Else
            Response.Write "File NOT found.<br>"
        End If

        ' 4. Read Text
        Response.Write "<h3>4. Reading Content...</h3>"
        content = fs.Read(path)
        Response.Write "<pre style='background:#eee; padding:10px;'>" & content & "</pre>"

        ' 5. File Size
        Response.Write "Size: " & fs.Size(path) & " bytes<br>"

        ' 6. List Files (Integration with your JSON/Array logic)
        Response.Write "<h3>6. List Files in root...</h3>"
        Dim files
        files = fs.List(".") ' List current dir
        
        For Each f In files
            Response.Write "File: " & f & "<br>"
        Next

        ' 7. Delete (Cleanup)
        ' Uncomment to test deletion
        ' fs.Delete(path)
    %>
    </div>
</body>
</html>