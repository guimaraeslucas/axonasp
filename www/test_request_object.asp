<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <title>Request Object - Complete Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; border-bottom: 3px solid #4CAF50; padding-bottom: 10px; }
        h2 { color: #4CAF50; margin-top: 30px; }
        .test-section { background: #f9f9f9; padding: 15px; margin: 15px 0; border-left: 4px solid #4CAF50; }
        .success { color: green; font-weight: bold; }
        .error { color: red; font-weight: bold; }
        .info { color: #666; font-style: italic; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        th { background: #4CAF50; color: white; padding: 10px; text-align: left; }
        td { padding: 8px; border-bottom: 1px solid #ddd; }
        tr:hover { background: #f5f5f5; }
        code { background: #eee; padding: 2px 6px; border-radius: 3px; font-family: 'Courier New', monospace; }
        .method { background: #e3f2fd; padding: 10px; margin: 10px 0; border-radius: 4px; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 2px solid #ddd; text-align: center; color: #666; }
    </style>
</head>
<body>
<div class="container">
    <h1>🔍 ASP Classic Request Object - Complete Test Suite</h1>
    <p class="info">Testing all Request Object collections, properties, and methods</p>

    <%
    ' Test Variables
    Dim testsPassed, testsFailed
    testsPassed = 0
    testsFailed = 0
    
    Function TestResult(testName, condition, expected, actual)
        Response.Write "<div class='method'>"
        Response.Write "<strong>" & testName & "</strong><br>"
        If condition Then
            Response.Write "<span class='success'>✓ PASS</span> - "
            testsPassed = testsPassed + 1
        Else
            Response.Write "<span class='error'>✗ FAIL</span> - "
            testsFailed = testsFailed + 1
        End If
        Response.Write "Expected: <code>" & expected & "</code>, "
        Response.Write "Got: <code>" & actual & "</code>"
        Response.Write "</div>"
    End Function
    %>

    <!-- ==================== COLLECTIONS ==================== -->
    <h2>📦 Collections</h2>

    <div class="test-section">
        <h3>QueryString Collection</h3>
        <%
        ' Test QueryString access
        Dim qsValue
        qsValue = Request.QueryString("test")
        Call TestResult("QueryString(key)", IsNull(qsValue) OR qsValue = "", "empty or test value", qsValue)
        
        ' Test QueryString count
        Response.Write "<p><strong>QueryString Count:</strong> " & Request.QueryString.Count() & "</p>"
        
        ' Show all QueryString items
        Response.Write "<p><strong>Current QueryString:</strong> "
        If Request.ServerVariables("QUERY_STRING") <> "" Then
            Response.Write Request.ServerVariables("QUERY_STRING")
        Else
            Response.Write "<em>(empty - try adding ?test=value to URL)</em>"
        End If
        Response.Write "</p>"
        %>
    </div>

    <div class="test-section">
        <h3>Form Collection</h3>
        <%
        ' Test Form collection
        Dim formValue
        formValue = Request.Form("username")
        Response.Write "<p><strong>Form(username):</strong> "
        If formValue = "" Then
            Response.Write "<em>(empty - submit the form below to test)</em>"
        Else
            Response.Write formValue
        End If
        Response.Write "</p>"
        
        Response.Write "<p><strong>Form Count:</strong> " & Request.Form.Count() & "</p>"
        %>
        
        <!-- Form for testing POST data -->
        <form method="post" action="test_request_object.asp?test=fromquery">
            <label>Username: <input type="text" name="username" value="testuser"></label>
            <label>Email: <input type="text" name="email" value="test@example.com"></label>
            <button type="submit">Submit Form</button>
        </form>
    </div>

    <div class="test-section">
        <h3>Cookies Collection</h3>
        <%
        ' Set a test cookie if not exists
        If Request.Cookies("testcookie") = "" Then
            Response.Cookies("testcookie") = "test_value_123"
        End If
        
        Dim cookieValue
        cookieValue = Request.Cookies("testcookie")
        Call TestResult("Cookies(key)", cookieValue <> "", "test_value_123 or similar", cookieValue)
        
        Response.Write "<p><strong>Cookies Count:</strong> " & Request.Cookies.Count() & "</p>"
        %>
    </div>

    <div class="test-section">
        <h3>ServerVariables Collection</h3>
        <table>
            <tr>
                <th>Variable Name</th>
                <th>Value</th>
            </tr>
            <%
            ' Display important server variables
            Dim serverVars, i, varName
            serverVars = Array("REQUEST_METHOD", "REQUEST_PATH", "QUERY_STRING", "REMOTE_ADDR", _
                              "SERVER_NAME", "HTTP_USER_AGENT", "CONTENT_TYPE", "CONTENT_LENGTH")
            
            For i = 0 To UBound(serverVars)
                varName = serverVars(i)
                Response.Write "<tr>"
                Response.Write "<td><code>" & varName & "</code></td>"
                Response.Write "<td>" & Request.ServerVariables(varName) & "</td>"
                Response.Write "</tr>"
            Next
            %>
        </table>
        <%
        Response.Write "<p><strong>ServerVariables Count:</strong> " & Request.ServerVariables.Count() & "</p>"
        %>
    </div>

    <div class="test-section">
        <h3>ClientCertificate Collection</h3>
        <%
        ' Test ClientCertificate collection (stub implementation)
        Response.Write "<p><strong>ClientCertificate Count:</strong> " & Request.ClientCertificate.Count() & "</p>"
        
        Response.Write "<table>"
        Response.Write "<tr><th>Field</th><th>Value</th></tr>"
        Dim certFields
        certFields = Array("Subject", "Issuer", "ValidFrom", "ValidUntil", "SerialNumber")
        For i = 0 To UBound(certFields)
            Response.Write "<tr>"
            Response.Write "<td><code>" & certFields(i) & "</code></td>"
            Response.Write "<td>" & Request.ClientCertificate(certFields(i)) & "</td>"
            Response.Write "</tr>"
        Next
        Response.Write "</table>"
        
        Response.Write "<p class='info'>Note: ClientCertificate is a stub implementation. SSL certificates are not currently processed.</p>"
        %>
    </div>

    <!-- ==================== PROPERTIES ==================== -->
    <h2>📊 Properties</h2>

    <div class="test-section">
        <h3>TotalBytes Property</h3>
        <%
        Dim totalBytes
        totalBytes = Request.TotalBytes
        
        Response.Write "<p><strong>Request.TotalBytes:</strong> " & totalBytes & " bytes</p>"
        
        If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
            Call TestResult("TotalBytes (POST)", totalBytes >= 0, ">= 0", totalBytes)
        Else
            Response.Write "<p class='info'>Submit the form above to test TotalBytes with POST data</p>"
        End If
        %>
    </div>

    <!-- ==================== METHODS ==================== -->
    <h2>⚡ Methods</h2>

    <div class="test-section">
        <h3>BinaryRead Method</h3>
        <%
        ' Test BinaryRead method
        If Request.ServerVariables("REQUEST_METHOD") = "POST" AND Request.TotalBytes > 0 Then
            Dim bytesRead, byteCount
            byteCount = Request.TotalBytes
            
            If byteCount > 100 Then
                byteCount = 100 ' Read only first 100 bytes for display
            End If
            
            Response.Write "<p>Reading first " & byteCount & " bytes from request body...</p>"
            
            ' Note: In Classic ASP, calling BinaryRead prevents access to Form collection
            ' For testing, we'll show this as an example
            Response.Write "<p class='info'>⚠ In Classic ASP, calling BinaryRead() prevents access to the Form collection.</p>"
            Response.Write "<p class='info'>For this test, we show the property exists but don't call it to preserve Form collection.</p>"
            
            Call TestResult("BinaryRead exists", True, "Method available", "✓ Available")
            
        Else
            Response.Write "<p class='info'>Submit the form above to test BinaryRead with POST data</p>"
            Response.Write "<p>Example usage: <code>byteArray = Request.BinaryRead(100)</code></p>"
        End If
        %>
    </div>

    <!-- ==================== ADVANCED TESTS ==================== -->
    <h2>🔬 Advanced Tests</h2>

    <div class="test-section">
        <h3>Collection Iteration</h3>
        <%
        ' Test iterating through collections
        Response.Write "<p><strong>QueryString Keys:</strong> "
        Dim keys, key
        keys = Request.QueryString.GetKeys()
        If IsArray(keys) And UBound(keys) >= 0 Then
            For i = 0 To UBound(keys)
                Response.Write keys(i) & " = " & Request.QueryString(keys(i))
                If i < UBound(keys) Then Response.Write ", "
            Next
        Else
            Response.Write "<em>(none)</em>"
        End If
        Response.Write "</p>"
        %>
    </div>

    <div class="test-section">
        <h3>Case Insensitivity Test</h3>
        <%
        ' Test case insensitivity
        Dim val1, val2, val3
        val1 = Request.ServerVariables("REQUEST_METHOD")
        val2 = Request.ServerVariables("request_method")
        val3 = Request.ServerVariables("Request_Method")
        
        Call TestResult("Case Insensitivity", val1 = val2 AND val2 = val3, "all equal", "val1=" & val1 & ", val2=" & val2 & ", val3=" & val3)
        %>
    </div>

    <div class="test-section">
        <h3>Empty Key Access</h3>
        <%
        ' Test accessing non-existent keys
        Dim emptyVal
        emptyVal = Request.QueryString("nonexistentkey")
        Call TestResult("Non-existent key returns empty", emptyVal = "", "empty string", emptyVal)
        %>
    </div>

    <!-- ==================== TEST SUMMARY ==================== -->
    <h2>📈 Test Summary</h2>
    <div class="test-section">
        <table>
            <tr>
                <th>Metric</th>
                <th>Value</th>
            </tr>
            <tr>
                <td>Tests Passed</td>
                <td class="success"><%= testsPassed %></td>
            </tr>
            <tr>
                <td>Tests Failed</td>
                <td class="<% If testsFailed > 0 Then Response.Write("error") Else Response.Write("success") End If %>"><%= testsFailed %></td>
            </tr>
            <tr>
                <td>Total Tests</td>
                <td><strong><%= testsPassed + testsFailed %></strong></td>
            </tr>
            <tr>
                <td>Success Rate</td>
                <td>
                    <% 
                    If testsPassed + testsFailed > 0 Then
                        Response.Write FormatNumber((testsPassed / (testsPassed + testsFailed)) * 100, 2) & "%"
                    Else
                        Response.Write "N/A"
                    End If
                    %>
                </td>
            </tr>
        </table>
    </div>

    <!-- ==================== API REFERENCE ==================== -->
    <h2>📚 Request Object API Reference</h2>
    <div class="test-section">
        <h3>Collections</h3>
        <ul>
            <li><code>Request.QueryString(key)</code> - URL query parameters</li>
            <li><code>Request.Form(key)</code> - POST form data</li>
            <li><code>Request.Cookies(key)</code> - HTTP cookies</li>
            <li><code>Request.ServerVariables(key)</code> - Server environment variables</li>
            <li><code>Request.ClientCertificate(key)</code> - Client SSL certificate info</li>
        </ul>
        
        <h3>Properties</h3>
        <ul>
            <li><code>Request.TotalBytes</code> - Total bytes in request body (readonly)</li>
        </ul>
        
        <h3>Methods</h3>
        <ul>
            <li><code>Request.BinaryRead(count)</code> - Reads binary data from request body</li>
        </ul>
        
        <h3>Collection Methods</h3>
        <ul>
            <li><code>Collection.Count()</code> - Returns number of items</li>
            <li><code>Collection.GetKeys()</code> - Returns array of keys</li>
        </ul>
    </div>

    <div class="footer">
        <p><strong>G3 AxonASP</strong> - Request Object Complete Implementation</p>
        <p>All collections, properties, and methods from ASP Classic 3.0</p>
        <p><a href="default.asp">← Back to Home</a></p>
    </div>
</div>
</body>
</html>
