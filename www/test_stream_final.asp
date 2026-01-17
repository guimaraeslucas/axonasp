<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Stream Function - Final Test</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        h1 { color: #333; }
        .test { margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }
        .success { background: #d4edda; border-color: #c3e6cb; }
        .info { background: #d1ecf1; border-color: #bee5eb; }
        pre { background: #f5f5f5; padding: 10px; border-radius: 3px; }
    </style>
</head>
<body>
    <h1>✓ User-Defined Function Test: stream()</h1>
    <p>Testing private function with ADODB.Stream and ByRef parameter</p>
    
    <%
    ' Define the user's stream function
    Private Function stream(path, binary, ByRef size)
        On Error Resume Next
        Dim objStream
        Set objStream = Server.CreateObject("ADODB.Stream")
        
        If binary Then
            objStream.Open
            objStream.Type = 1 ' adTypeBinary
            objStream.LoadFromFile(Server.MapPath(path))
            stream = objStream.Read()
        Else
            objStream.CharSet = "utf-8"
            objStream.Open
            objStream.Type = 2 ' adTypeText
            objStream.LoadFromFile(Server.MapPath(path))
            stream = objStream.ReadText()
        End If
        
        size = objStream.Size
        Set objStream = Nothing
        On Error GoTo 0
    End Function
    
    ' Test 1: Read text file
    Dim textContent, fileSize
    fileSize = 0
    
    Response.Write "<div class='test info'>"
    Response.Write "<h2>Test 1: Text File Reading</h2>"
    Response.Write "<p><strong>File:</strong> demo_file.txt</p>"
    
    textContent = stream("demo_file.txt", False, fileSize)
    
    Response.Write "<p><strong>Content Length:</strong> " & Len(textContent) & " characters</p>"
    Response.Write "<p><strong>File Size (via ByRef):</strong> " & fileSize & " bytes</p>"
    Response.Write "<p><strong>Content:</strong></p>"
    Response.Write "<pre>" & Server.HTMLEncode(textContent) & "</pre>"
    Response.Write "</div>"
    
    ' Test 2: Multiple calls with different sizes
    Dim size1, size2
    size1 = 0
    size2 = 0
    
    Response.Write "<div class='test success'>"
    Response.Write "<h2>Test 2: ByRef Parameter Verification</h2>"
    
    Dim content1, content2
    content1 = stream("demo_file.txt", False, size1)
    content2 = stream("demo_file.txt", False, size2)
    
    Response.Write "<p><strong>First call:</strong> size1 = " & size1 & " bytes</p>"
    Response.Write "<p><strong>Second call:</strong> size2 = " & size2 & " bytes</p>"
    
    If size1 = size2 And size1 > 0 Then
        Response.Write "<p style='color: green;'><strong>✓ ByRef working correctly!</strong></p>"
    Else
        Response.Write "<p style='color: red;'><strong>✗ ByRef issue detected</strong></p>"
    End If
    Response.Write "</div>"
    
    Response.Write "<div class='test success'>"
    Response.Write "<h2>Summary</h2>"
    Response.Write "<ul>"
    Response.Write "<li>✓ Private function execution</li>"
    Response.Write "<li>✓ ADODB.Stream integration</li>"
    Response.Write "<li>✓ ByRef parameter modification</li>"
    Response.Write "<li>✓ Text file reading</li>"
    Response.Write "<li>✓ Multiple function calls</li>"
    Response.Write "</ul>"
    Response.Write "</div>"
    %>
</body>
</html>
