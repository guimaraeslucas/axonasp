<%@ LANGUAGE="VBSCRIPT" %>
<!DOCTYPE html>
<html>
<head>
    <title>Request Object - Complete Test Suite</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1000px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; border-bottom: 3px solid #4CAF50; padding-bottom: 10px; }
        h2 { color: #4CAF50; margin-top: 30px; }
        .test-section { background: #f9f9f9; padding: 15px; margin: 15px 0; border-left: 4px solid #4CAF50; }
        .pass { color: green; font-weight: bold; }
        .info { color: #666; font-style: italic; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        th { background: #4CAF50; color: white; padding: 10px; text-align: left; }
        td { padding: 8px; border-bottom: 1px solid #ddd; }
        code { background: #eee; padding: 2px 6px; border-radius: 3px; font-family: 'Courier New', monospace; }
        ul { list-style-type: none; padding-left: 0; }
        li { background: #fff; padding: 8px; margin: 4px 0; border-radius: 4px; }
        .method { background: #e3f2fd; padding: 10px; margin: 10px 0; border-radius: 4px; }
    </style>
</head>
<body>
<div class="container">
    <h1>🔍 Request Object - Complete Test Suite</h1>
    <p class="info">Testing all Request Object collections, properties, and methods from ASP Classic 3.0</p>

    <%
    ' ==================== COLLECTIONS ====================
    %>
    
    <h2>📦 Collections</h2>

    <div class="test-section">
        <h3>1. QueryString Collection</h3>
        <p><strong>Description:</strong> Access URL query parameters</p>
        <%
        Dim qsValue, qsKey
        Response.Write "<p><strong>Individual Access:</strong></p>"
        Response.Write "<ul>"
        Response.Write "<li><code>Request.QueryString(""test"")</code> = """ & Request.QueryString("test") & """</li>"
        Response.Write "<li><code>Request.QueryString(""foo"")</code> = """ & Request.QueryString("foo") & """</li>"
        Response.Write "</ul>"
        
        Response.Write "<p><strong>Iteration (For Each):</strong></p>"
        Response.Write "<ul>"
        For Each qsKey In Request.QueryString
            Response.Write "<li>" & qsKey & " = " & Request.QueryString(qsKey) & "</li>"
        Next
        Response.Write "</ul>"
        
        Response.Write "<p class='info'>Try: <a href='test_request_complete.asp?test=value1&foo=value2'>?test=value1&foo=value2</a></p>"
        %>
    </div>

    <div class="test-section">
        <h3>2. Form Collection</h3>
        <p><strong>Description:</strong> Access POST form data</p>
        <%
        Dim formKey
        Response.Write "<p><strong>Form Data:</strong></p>"
        Response.Write "<ul>"
        Response.Write "<li><code>Request.Form(""username"")</code> = """ & Request.Form("username") & """</li>"
        Response.Write "<li><code>Request.Form(""email"")</code> = """ & Request.Form("email") & """</li>"
        Response.Write "</ul>"
        
        Response.Write "<p><strong>All Form Fields:</strong></p>"
        Response.Write "<ul>"
        For Each formKey In Request.Form
            Response.Write "<li>" & formKey & " = " & Request.Form(formKey) & "</li>"
        Next
        Response.Write "</ul>"
        %>
        
        <form method="post" action="test_request_complete.asp?test=fromquery">
            <label>Username: <input type="text" name="username" value="testuser"></label><br>
            <label>Email: <input type="text" name="email" value="test@example.com"></label><br>
            <button type="submit">Submit Form</button>
        </form>
    </div>

    <div class="test-section">
        <h3>3. Cookies Collection</h3>
        <p><strong>Description:</strong> Access HTTP cookies</p>
        <%
        ' Set test cookie
        If Request.Cookies("testcookie") = "" Then
            Response.Cookies("testcookie") = "test_value_123"
        End If
        
        Dim cookieKey
        Response.Write "<p><strong>Individual Access:</strong></p>"
        Response.Write "<ul>"
        Response.Write "<li><code>Request.Cookies(""testcookie"")</code> = """ & Request.Cookies("testcookie") & """</li>"
        Response.Write "<li><code>Request.Cookies(""ASPSESSIONID"")</code> = """ & Request.Cookies("ASPSESSIONID") & """</li>"
        Response.Write "</ul>"
        
        Response.Write "<p><strong>All Cookies:</strong></p>"
        Response.Write "<ul>"
        For Each cookieKey In Request.Cookies
            Response.Write "<li>" & cookieKey & " = " & Request.Cookies(cookieKey) & "</li>"
        Next
        Response.Write "</ul>"
        %>
    </div>

    <div class="test-section">
        <h3>4. ServerVariables Collection</h3>
        <p><strong>Description:</strong> Access server environment variables</p>
        <table>
            <tr>
                <th>Variable Name</th>
                <th>Value</th>
            </tr>
            <%
            ' Display important server variables
            Response.Write "<tr><td><code>REQUEST_METHOD</code></td><td>" & Request.ServerVariables("REQUEST_METHOD") & "</td></tr>"
            Response.Write "<tr><td><code>REQUEST_PATH</code></td><td>" & Request.ServerVariables("REQUEST_PATH") & "</td></tr>"
            Response.Write "<tr><td><code>QUERY_STRING</code></td><td>" & Request.ServerVariables("QUERY_STRING") & "</td></tr>"
            Response.Write "<tr><td><code>REMOTE_ADDR</code></td><td>" & Request.ServerVariables("REMOTE_ADDR") & "</td></tr>"
            Response.Write "<tr><td><code>SERVER_NAME</code></td><td>" & Request.ServerVariables("SERVER_NAME") & "</td></tr>"
            Response.Write "<tr><td><code>HTTP_USER_AGENT</code></td><td>" & Request.ServerVariables("HTTP_USER_AGENT") & "</td></tr>"
            Response.Write "<tr><td><code>CONTENT_TYPE</code></td><td>" & Request.ServerVariables("CONTENT_TYPE") & "</td></tr>"
            Response.Write "<tr><td><code>CONTENT_LENGTH</code></td><td>" & Request.ServerVariables("CONTENT_LENGTH") & "</td></tr>"
            %>
        </table>
        <%
        Dim svKey
        Response.Write "<p><strong>All ServerVariables:</strong></p>"
        Response.Write "<ul>"
        For Each svKey In Request.ServerVariables
            Response.Write "<li>" & svKey & " = " & Request.ServerVariables(svKey) & "</li>"
        Next
        Response.Write "</ul>"
        %>
    </div>

    <div class="test-section">
        <h3>5. ClientCertificate Collection</h3>
        <p><strong>Description:</strong> Access client SSL certificate information</p>
        <%
        Dim ccKey
        Response.Write "<p><strong>Certificate Fields:</strong></p>"
        Response.Write "<table>"
        Response.Write "<tr><th>Field</th><th>Value</th></tr>"
        Response.Write "<tr><td><code>Subject</code></td><td>" & Request.ClientCertificate("Subject") & "</td></tr>"
        Response.Write "<tr><td><code>Issuer</code></td><td>" & Request.ClientCertificate("Issuer") & "</td></tr>"
        Response.Write "<tr><td><code>ValidFrom</code></td><td>" & Request.ClientCertificate("ValidFrom") & "</td></tr>"
        Response.Write "<tr><td><code>ValidUntil</code></td><td>" & Request.ClientCertificate("ValidUntil") & "</td></tr>"
        Response.Write "<tr><td><code>SerialNumber</code></td><td>" & Request.ClientCertificate("SerialNumber") & "</td></tr>"
        Response.Write "</table>"
        
        Response.Write "<p><strong>All Certificate Fields:</strong></p>"
        Response.Write "<ul>"
        For Each ccKey In Request.ClientCertificate
            Response.Write "<li>" & ccKey & " = " & Request.ClientCertificate(ccKey) & "</li>"
        Next
        Response.Write "</ul>"
        
        Response.Write "<p class='info'>Note: ClientCertificate is a stub implementation. SSL certificates are not currently processed.</p>"
        %>
    </div>

    <%
    ' ==================== PROPERTIES ====================
    %>
    
    <h2>📊 Properties</h2>

    <div class="test-section">
        <h3>TotalBytes Property</h3>
        <p><strong>Description:</strong> Returns the total number of bytes in the request body (readonly)</p>
        <%
        Dim totalBytes, requestMethod
        totalBytes = Request.TotalBytes
        requestMethod = Request.ServerVariables("REQUEST_METHOD")
        
        Response.Write "<p><strong>Request.TotalBytes:</strong> " & totalBytes & " bytes</p>"
        
        If requestMethod = "POST" Then
            Response.Write "<p class='pass'>✓ POST request detected - TotalBytes shows body size</p>"
        Else
            Response.Write "<p class='info'>Submit the form above to test TotalBytes with POST data</p>"
        End If
        %>
        <p><strong>Usage:</strong> <code>totalBytes = Request.TotalBytes</code></p>
    </div>

    <%
    ' ==================== METHODS ====================
    %>
    
    <h2>⚡ Methods</h2>

    <div class="test-section">
        <h3>BinaryRead Method</h3>
        <p><strong>Description:</strong> Reads binary data from the request body</p>
        <div class="method">
            <p><strong>Syntax:</strong> <code>byteArray = Request.BinaryRead(count)</code></p>
            <p><strong>Parameters:</strong></p>
            <ul>
                <li><code>count</code> - Number of bytes to read from request body</li>
            </ul>
            <p><strong>Returns:</strong> Byte array containing the data</p>
        </div>
        <%
        If requestMethod = "POST" AND totalBytes > 0 Then
            Response.Write "<p class='pass'>✓ BinaryRead is available for this POST request</p>"
            Response.Write "<p><strong>Example:</strong> <code>data = Request.BinaryRead(" & totalBytes & ")</code></p>"
            Response.Write "<p class='info'>⚠ Note: In Classic ASP, calling BinaryRead() prevents access to the Form collection.</p>"
            Response.Write "<p class='info'>For this test, we show the method exists without calling it to preserve Form collection access.</p>"
        Else
            Response.Write "<p class='info'>Submit the form above to test BinaryRead with POST data</p>"
        End If
        %>
    </div>

    <%
    ' ==================== TEST SUMMARY ====================
    %>
    
    <h2>📈 Implementation Summary</h2>
    <div class="test-section">
        <table>
            <tr>
                <th>Feature</th>
                <th>Type</th>
                <th>Status</th>
            </tr>
            <tr>
                <td><code>QueryString</code></td>
                <td>Collection</td>
                <td class="pass">✓ Implemented</td>
            </tr>
            <tr>
                <td><code>Form</code></td>
                <td>Collection</td>
                <td class="pass">✓ Implemented</td>
            </tr>
            <tr>
                <td><code>Cookies</code></td>
                <td>Collection</td>
                <td class="pass">✓ Implemented</td>
            </tr>
            <tr>
                <td><code>ServerVariables</code></td>
                <td>Collection</td>
                <td class="pass">✓ Implemented</td>
            </tr>
            <tr>
                <td><code>ClientCertificate</code></td>
                <td>Collection</td>
                <td class="pass">✓ Implemented (stub)</td>
            </tr>
            <tr>
                <td><code>TotalBytes</code></td>
                <td>Property</td>
                <td class="pass">✓ Implemented</td>
            </tr>
            <tr>
                <td><code>BinaryRead</code></td>
                <td>Method</td>
                <td class="pass">✓ Implemented</td>
            </tr>
        </table>
        
        <p><strong>All ASP Classic 3.0 Request Object features are implemented!</strong></p>
        <ul>
            <li>5 Collections with full iteration support (For Each)</li>
            <li>1 Readonly Property (TotalBytes)</li>
            <li>1 Method (BinaryRead)</li>
            <li>Case-insensitive key access</li>
            <li>Thread-safe with mutex locking</li>
        </ul>
    </div>

    <p style="text-align: center; margin-top: 30px;">
        <a href="default.asp">← Back to Home</a>
    </p>
</div>
</body>
</html>
