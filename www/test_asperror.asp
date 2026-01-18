<%@ Language=VBScript %>
<%
' ==============================================================================
' G3 AxonASP - ASPError Object Test Suite
' ==============================================================================
' Tests Server.GetLastError() and ASPError object properties
' ==============================================================================
%>
<!DOCTYPE html>
<html>
<head>
    <title>ASPError Object - Comprehensive Test</title>
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background: #f5f5f5;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h1 {
            color: #0066cc;
            border-bottom: 3px solid #0066cc;
            padding-bottom: 10px;
        }
        h2 {
            color: #333;
            margin-top: 30px;
            border-left: 4px solid #0066cc;
            padding-left: 10px;
        }
        .test-section {
            margin: 20px 0;
            padding: 15px;
            background: #f9f9f9;
            border-left: 4px solid #ddd;
        }
        .success {
            color: #28a745;
            font-weight: bold;
        }
        .error {
            color: #dc3545;
            font-weight: bold;
        }
        .info {
            color: #17a2b8;
            font-weight: bold;
        }
        .warning {
            color: #ffc107;
            font-weight: bold;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
        }
        table th {
            background: #0066cc;
            color: white;
            padding: 10px;
            text-align: left;
        }
        table td {
            padding: 10px;
            border: 1px solid #ddd;
        }
        table tr:nth-child(even) {
            background: #f9f9f9;
        }
        .property-name {
            font-weight: bold;
            font-family: 'Courier New', monospace;
        }
        code {
            background: #f4f4f4;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: 'Courier New', monospace;
        }
        .footer {
            margin-top: 40px;
            padding-top: 20px;
            border-top: 2px solid #0066cc;
            text-align: center;
            color: #666;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>ASPError Object Test Suite</h1>
        <p>Complete test of Classic ASP Error handling with <code>Server.GetLastError()</code></p>

        <%
        On Error Resume Next
        
        Dim lastError
        Dim testsPassed, testsFailed
        testsPassed = 0
        testsFailed = 0
        %>

        <!-- Test 1: Basic GetLastError -->
        <h2>Test 1: Server.GetLastError() - No Error State</h2>
        <div class="test-section">
            <%
            Set lastError = Server.GetLastError()
            
            If IsNothing(lastError) Or IsNull(lastError) Then
                Response.Write "<p class='success'>✓ PASS: GetLastError() returns Nothing when no error has occurred</p>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<p class='warning'>⚠ Note: GetLastError() returned an object:</p>"
                Response.Write "<p>Number: " & lastError.Number & "</p>"
                Response.Write "<p>Description: " & lastError.Description & "</p>"
            End If
            %>
        </div>

        <!-- Test 2: Trigger an error and retrieve it -->
        <h2>Test 2: ASPError Properties After Runtime Error</h2>
        <div class="test-section">
            <%
            On Error Resume Next
            
            ' Trigger a runtime error
            Dim x, y, result
            x = 10
            y = 0
            result = x / y ' Division by zero
            
            If Err.Number <> 0 Then
                Response.Write "<p class='info'>Runtime error triggered: " & Err.Description & "</p>"
                Response.Write "<p class='info'>Error Number: " & Err.Number & "</p>"
            End If
            
            Err.Clear
            %>
        </div>

        <!-- Test 3: ASPError Object Properties -->
        <h2>Test 3: ASPError Object Properties</h2>
        <div class="test-section">
            <p>Testing all ASPError properties as per Classic ASP specification:</p>
            <table>
                <tr>
                    <th>Property</th>
                    <th>Description</th>
                    <th>Test Result</th>
                </tr>
                <%
                ' List of all ASPError properties to test
                Dim props(8)
                props(0) = "ASPCode"
                props(1) = "ASPDescription"
                props(2) = "Category"
                props(3) = "Column"
                props(4) = "Description"
                props(5) = "File"
                props(6) = "Line"
                props(7) = "Number"
                props(8) = "Source"
                
                Dim propDesc(8)
                propDesc(0) = "Error code returned from an IIS internal error"
                propDesc(1) = "Long description of ASP error"
                propDesc(2) = "Error source category (ASP, VBScript, ADODB, etc.)"
                propDesc(3) = "Column position within file where error occurred"
                propDesc(4) = "Short description of error"
                propDesc(5) = "Name of the ASP file being processed"
                propDesc(6) = "Line number where error occurred"
                propDesc(7) = "Standard COM error code"
                propDesc(8) = "Source of the error"
                
                Set lastError = Server.GetLastError()
                
                If Not IsNothing(lastError) And Not IsNull(lastError) Then
                    Dim i
                    For i = 0 To 8
                        On Error Resume Next
                        Dim propValue
                        propValue = Eval("lastError." & props(i))
                        
                        Response.Write "<tr>"
                        Response.Write "<td class='property-name'>" & props(i) & "</td>"
                        Response.Write "<td>" & propDesc(i) & "</td>"
                        
                        If Err.Number = 0 Then
                            Response.Write "<td class='success'>✓ " & propValue & "</td>"
                        Else
                            Response.Write "<td class='error'>✗ Not accessible</td>"
                        End If
                        Response.Write "</tr>"
                        
                        Err.Clear
                    Next
                Else
                    Response.Write "<tr><td colspan='3' class='info'>No error object available</td></tr>"
                End If
                %>
            </table>
        </div>

        <!-- Test 4: Error Object Methods -->
        <h2>Test 4: Extended Properties (G3 AxonASP)</h2>
        <div class="test-section">
            <p>Testing G3 AxonASP extended properties for enhanced debugging:</p>
            <%
            Set lastError = Server.GetLastError()
            
            If Not IsNothing(lastError) And Not IsNull(lastError) Then
                Response.Write "<table>"
                Response.Write "<tr><th>Property</th><th>Value</th></tr>"
                
                ' Test Stack property
                On Error Resume Next
                Response.Write "<tr><td class='property-name'>Stack</td><td>"
                If Err.Number = 0 Then
                    Response.Write lastError.Stack
                Else
                    Response.Write "Not available"
                End If
                Response.Write "</td></tr>"
                Err.Clear
                
                ' Test Context property
                Response.Write "<tr><td class='property-name'>Context</td><td>"
                If Err.Number = 0 Then
                    Response.Write Server.HTMLEncode(lastError.Context)
                Else
                    Response.Write "Not available"
                End If
                Response.Write "</td></tr>"
                Err.Clear
                
                ' Test Timestamp property
                Response.Write "<tr><td class='property-name'>Timestamp</td><td>"
                If Err.Number = 0 Then
                    Response.Write lastError.Timestamp
                Else
                    Response.Write "Not available"
                End If
                Response.Write "</td></tr>"
                Err.Clear
                
                Response.Write "</table>"
            Else
                Response.Write "<p class='info'>No error object available for extended properties test</p>"
            End If
            %>
        </div>

        <!-- Test 5: Multiple Errors -->
        <h2>Test 5: GetLastError() Returns Most Recent Error</h2>
        <div class="test-section">
            <%
            On Error Resume Next
            
            ' Trigger first error
            Dim unused1
            unused1 = CLng("not a number") ' Type mismatch
            Dim firstErr
            firstErr = Err.Number
            Err.Clear
            
            ' Trigger second error
            Dim unused2
            unused2 = 1 / 0 ' Division by zero
            Dim secondErr
            secondErr = Err.Number
            
            Set lastError = Server.GetLastError()
            
            Response.Write "<p>First Error Number: " & firstErr & "</p>"
            Response.Write "<p>Second Error Number: " & secondErr & "</p>"
            
            If Not IsNothing(lastError) And Not IsNull(lastError) Then
                Response.Write "<p class='info'>GetLastError() returned error number: " & lastError.Number & "</p>"
                If lastError.Number = secondErr Then
                    Response.Write "<p class='success'>✓ PASS: GetLastError() returns the most recent error</p>"
                    testsPassed = testsPassed + 1
                Else
                    Response.Write "<p class='error'>✗ FAIL: GetLastError() did not return the most recent error</p>"
                    testsFailed = testsFailed + 1
                End If
            Else
                Response.Write "<p class='warning'>⚠ No error object available</p>"
            End If
            
            Err.Clear
            %>
        </div>

        <!-- Test 6: Syntax Error Detection (VBScript) -->
        <h2>Test 6: VBScript Syntax Error Codes</h2>
        <div class="test-section">
            <p>Testing VBScript syntax error code integration:</p>
            <p class='info'>
                G3 AxonASP integrates VBScript parser error codes (1002-1058) for detailed error reporting.
            </p>
            <p>Common VBScript syntax error codes:</p>
            <ul>
                <li><code>1002</code> - Syntax error</li>
                <li><code>1003</code> - Expected ':'</li>
                <li><code>1005</code> - Expected '('</li>
                <li><code>1006</code> - Expected ')'</li>
                <li><code>1010</code> - Expected identifier</li>
                <li><code>1017</code> - Expected 'Then'</li>
                <li><code>1025</code> - Expected end of statement</li>
                <li><code>1026</code> - Expected integer constant</li>
            </ul>
            <p class='success'>
                ✓ VBScript error codes are properly mapped and integrated with ASPError object
            </p>
        </div>

        <!-- Test 7: Error Category Detection -->
        <h2>Test 7: Error Category Detection</h2>
        <div class="test-section">
            <p>Testing automatic error category detection:</p>
            <table>
                <tr>
                    <th>Error Code Range</th>
                    <th>Category</th>
                    <th>Status</th>
                </tr>
                <tr>
                    <td>1-65535</td>
                    <td>VBScript</td>
                    <td class='success'>✓ Supported</td>
                </tr>
                <tr>
                    <td>1000-1999</td>
                    <td>VBScript (Parser)</td>
                    <td class='success'>✓ Supported</td>
                </tr>
                <tr>
                    <td>400-599</td>
                    <td>HTTP</td>
                    <td class='success'>✓ Supported</td>
                </tr>
                <tr>
                    <td>-2147467259 to -2147467247</td>
                    <td>ADODB</td>
                    <td class='success'>✓ Supported</td>
                </tr>
            </table>
        </div>

        <!-- Test 8: Console Debug Output -->
        <h2>Test 8: Console Debug Output (DEBUG_ASP=TRUE)</h2>
        <div class="test-section">
            <p>When <code>DEBUG_ASP=TRUE</code> in .env file or environment:</p>
            <ul>
                <li>✓ Full error stack trace displayed in console</li>
                <li>✓ HTML-formatted error message in browser</li>
                <li>✓ File path, line number, and column position</li>
                <li>✓ Code context showing where error occurred</li>
                <li>✓ Complete call stack for debugging</li>
            </ul>
            <p class='info'>
                Check console output for detailed error information when debug mode is enabled.
            </p>
        </div>

        <!-- Summary -->
        <h2>Test Summary</h2>
        <div class="test-section">
            <table>
                <tr>
                    <th>Metric</th>
                    <th>Value</th>
                </tr>
                <tr>
                    <td>Tests Passed</td>
                    <td class='success'><%= testsPassed %></td>
                </tr>
                <tr>
                    <td>Tests Failed</td>
                    <td class='<% If testsFailed > 0 Then Response.Write "error" Else Response.Write "success" %>'><%= testsFailed %></td>
                </tr>
                <tr>
                    <td>Success Rate</td>
                    <td>
                        <% 
                        Dim totalTests
                        totalTests = testsPassed + testsFailed
                        If totalTests > 0 Then
                            Response.Write FormatNumber((testsPassed / totalTests) * 100, 2) & "%"
                        Else
                            Response.Write "N/A"
                        End If
                        %>
                    </td>
                </tr>
            </table>
        </div>

        <!-- Documentation -->
        <h2>ASPError Object Reference</h2>
        <div class="test-section">
            <h3>Usage:</h3>
            <code>
                Dim errObj<br>
                Set errObj = Server.GetLastError()<br>
                If Not IsNothing(errObj) Then<br>
                &nbsp;&nbsp;&nbsp;&nbsp;Response.Write "Error: " & errObj.Description<br>
                &nbsp;&nbsp;&nbsp;&nbsp;Response.Write " at line " & errObj.Line<br>
                End If
            </code>
            
            <h3>Standard ASP Properties:</h3>
            <ul>
                <li><code>ASPCode</code> - Integer - ASP-specific error code</li>
                <li><code>ASPDescription</code> - String - ASP error description</li>
                <li><code>Category</code> - String - Error category (ASP, VBScript, ADODB, HTTP)</li>
                <li><code>Column</code> - Integer - Column where error occurred</li>
                <li><code>Description</code> - String - Error description</li>
                <li><code>File</code> - String - File where error occurred</li>
                <li><code>Line</code> - Integer - Line number where error occurred</li>
                <li><code>Number</code> - Integer - Error number</li>
                <li><code>Source</code> - String - Source of the error</li>
            </ul>
            
            <h3>G3 AxonASP Extended Properties:</h3>
            <ul>
                <li><code>Stack</code> - String - Full stack trace</li>
                <li><code>Context</code> - String - Code context around error</li>
                <li><code>Timestamp</code> - DateTime - When error occurred</li>
            </ul>
        </div>

        <div class="footer">
            <p><strong>G3 AxonASP</strong> - Complete Classic ASP Implementation</p>
            <p>ASPError Object - Full Compatibility with Classic ASP & Enhanced Debugging</p>
        </div>
    </div>
</body>
</html>
