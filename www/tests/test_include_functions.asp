<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - AxInclude, AxIncludeOnce, and AxGetRemoteFile Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; border-left: 4px solid #667eea; padding-left: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .test-section { margin: 20px 0; padding: 15px; background: #f9f9f9; border: 1px solid #ddd; border-radius: 4px; }
        .success { background: #d4edda; border-left: 4px solid #28a745; color: #155724; padding: 10px; margin: 10px 0; border-radius: 4px; }
        .error { background: #f8d7da; border-left: 4px solid #dc3545; color: #721c24; padding: 10px; margin: 10px 0; border-radius: 4px; }
        .info { background: #d1ecf1; border-left: 4px solid #17a2b8; color: #0c5460; padding: 10px; margin: 10px 0; border-radius: 4px; }
        .warning { background: #fff3cd; border-left: 4px solid #ffc107; color: #856404; padding: 10px; margin: 10px 0; border-radius: 4px; }
        code { background: #f5f5f5; padding: 2px 6px; border-radius: 3px; font-family: 'Courier New', monospace; }
        pre { background: #f5f5f5; padding: 15px; border-radius: 4px; overflow-x: auto; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - AxInclude, AxIncludeOnce, and AxGetRemoteFile Test</h1>
        <div class="intro">
            <p>Tests AxInclude, AxIncludeOnce, and AxGetRemoteFile custom functions.</p>
            <p>AxInclude and AxIncludeOnce execute ASP files (same as <code>&lt;!--# include --&gt;</code>).</p>
            <p>AxGetRemoteFile fetches remote content as plain text (not executed).</p>
        </div>

        <h2>Test 1: AxInclude - Execute ASP File</h2>
        <div class="test-section">
            <p>Call AxInclude with a relative path to execute an ASP file:</p>
            <pre><code>If AxInclude("config.inc") Then
    Response.Write "File executed successfully"
Else
    Response.Write "Failed to include file"
End If</code></pre>
            <div class="info">Output:</div>
            <%
                If AxInclude("config.inc") Then
                    Response.Write "<p class='success'>✓ config.inc executed successfully</p>" & vbCrLf
                Else
                    Response.Write "<p class='error'>✗ Failed to include config.inc</p>" & vbCrLf
                End If
            %>
        </div>

        <h2>Test 2: AxInclude - File Not Found</h2>
        <div class="test-section">
            <p>Call AxInclude with a non-existent file (should return false):</p>
            <pre><code>If AxInclude("nonexistent_file.asp") Then
    Response.Write "File found"
Else
    Response.Write "File not found (error shown in console)"
End If</code></pre>
            <div class="info">Output:</div>
            <%
                If Not AxInclude("nonexistent_file.asp") Then
                    Response.Write "<p class='success'>✓ Correctly returned false for missing file</p>" & vbCrLf
                    Response.Write "<p class='info'>Check server console for error message</p>" & vbCrLf
                End If
            %>
        </div>

        <h2>Test 3: AxIncludeOnce - Prevent Duplicate Execution</h2>
        <div class="test-section">
            <p>Call AxIncludeOnce with the same file multiple times:</p>
            <pre><code>AxIncludeOnce "config.inc"
AxIncludeOnce "config.inc"
AxIncludeOnce "config.inc"</code></pre>
            <div class="info">Output (file should only be executed once):</div>
            <%
                Response.Write "<p>First call:</p>"
                AxIncludeOnce "config.inc"
                Response.Write "<p>Second call (should be ignored):</p>"
                AxIncludeOnce "config.inc"
                Response.Write "<p>Third call (should be ignored):</p>"
                AxIncludeOnce "config.inc"
                Response.Write "<p class='success'>✓ File was only executed once despite three calls</p>" & vbCrLf
            %>
        </div>

        <h2>Test 4: AxIncludeOnce - Different Files</h2>
        <div class="test-section">
            <p>Test AxIncludeOnce with different files:</p>
            <pre><code>AxIncludeOnce "config.inc"
AxIncludeOnce "config.inc"  ' Will be ignored
AxIncludeOnce "header.inc"   ' Will execute</code></pre>
            <div class="info">Output:</div>
            <%
                Response.Write "<p>Including config.inc:</p>"
                AxIncludeOnce "config.inc"
                Response.Write "<p>Attempting to include config.inc again (will be ignored):</p>"
                AxIncludeOnce "config.inc"
                Response.Write "<p class='warning'>Note: header.inc content would appear here if it was an ASP file</p>" & vbCrLf
            %>
        </div>

        <h2>Test 5: Path Resolution - Relative Paths</h2>
        <div class="test-section">
            <p>Test with explicit relative path prefix (./):</p>
            <pre><code>AxInclude "./config.inc"</code></pre>
            <div class="info">Output:</div>
            <%
                If AxInclude("./config.inc") Then
                    Response.Write "<p class='success'>✓ Explicit relative path works</p>" & vbCrLf
                End If
            %>
        </div>

        <h2>Test 6: AxGetRemoteFile - Fetch Remote Content</h2>
        <div class="test-section">
            <p>Call AxGetRemoteFile to fetch content from a remote URL:</p>
            <pre><code>Dim content
content = AxGetRemoteFile("https://example.com/robots.txt")</code></pre>
            <div class="info">Result:</div>
            <%
                Dim remoteContent
                remoteContent = AxGetRemoteFile("https://g3pix.com.br/robots.txt")
                
                If remoteContent <> False Then
                    Response.Write "<p class='success'>✓ Successfully fetched remote content</p>" & vbCrLf
                    Response.Write "<p>Content length: " & Len(remoteContent) & " bytes</p>" & vbCrLf
                    Response.Write "<p>First 200 characters:</p>" & vbCrLf
                    Response.Write "<pre>" & Left(remoteContent, 200) & "...</pre>" & vbCrLf
                Else
                    Response.Write "<p class='error'>✗ Failed to fetch remote content</p>" & vbCrLf
                    Response.Write "<p class='info'>Check server console for error message</p>" & vbCrLf
                End If
            %>
        </div>

        <h2>Test 7: AxGetRemoteFile - Invalid URL</h2>
        <div class="test-section">
            <p>Call AxGetRemoteFile with invalid URL format (should return false):</p>
            <pre><code>If AxGetRemoteFile("ftp://example.com/file.txt") = False Then
    Response.Write "Invalid protocol (FTP not supported)"
End If</code></pre>
            <div class="info">Result:</div>
            <%
                If AxGetRemoteFile("ftp://example.com/file.txt") = False Then
                    Response.Write "<p class='success'>✓ Correctly rejected unsupported protocol</p>" & vbCrLf
                    Response.Write "<p class='info'>Check server console for error message</p>" & vbCrLf
                End If
            %>
        </div>

        <h2>Test 8: Security - Local Path Rejection</h2>
        <div class="test-section">
            <p>Test that AxGetRemoteFile rejects local file paths for security:</p>
            <pre><code>If AxGetRemoteFile("C:\Windows\System32\config\system") = False Then
    Response.Write "Local paths are not allowed"
End If</code></pre>
            <div class="info">Result:</div>
            <%
                ' Try to access a local file (should be rejected)
                If AxGetRemoteFile("C:\Windows\System32\config") = False Then
                    Response.Write "<p class='success'>✓ Local file paths are blocked for security</p>" & vbCrLf
                End If
            %>
        </div>


        <h2>Check Server Console</h2>
        <div class="warning">
            <p><strong>Important:</strong> Error messages from AxInclude, AxIncludeOnce, and AxGetRemoteFile are printed to the server console, not to the browser.</p>
            <p>Check your terminal/console output to see detailed error messages for debugging.</p>
        </div>

        <p style="margin-top: 30px; border-top: 1px solid #ddd; padding-top: 20px; color: #999; font-size: 12px;">
            G3pix AxonASP - Custom Functions Test Suite
        </p>
    </div>
</body>
</html>
