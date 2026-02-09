<% @Language=VBScript %>
<%
' AxonASP WScript.Shell and AxExecute Testing
' This file tests both the WScript.Shell object and the AxExecute function
'
' To run tests, visit: http://localhost:4050/tests/test_wscript_shell.asp
' Or uncomment specific tests below

Option Compare Binary
Response.ContentType = "text/html"

Dim testsPassed, testsFailed, testResults
testsPassed = 0
testsFailed = 0
testResults = ""

' Helper function to log test results
Sub AssertTrue(expression, testName)
    If expression Then
        testsPassed = testsPassed + 1
        testResults = testResults & "<div style='color: green;'>[PASS] " & testName & "</div>"
    Else
        testsFailed = testsFailed + 1
        testResults = testResults & "<div style='color: red;'>[FAIL] " & testName & "</div>"
    End If
End Sub

Function AssertEqual(actual, expected, testName)
    If actual = expected Then
        testsPassed = testsPassed + 1
        testResults = testResults & "<div style='color: green;'>[PASS] " & testName & "</div>"
    Else
        testsFailed = testsFailed + 1
        testResults = testResults & "<div style='color: red;'>[FAIL] " & testName & " (expected: " & expected & ", got: " & actual & ")</div>"
    End If
End Function

Function AssertNotNull(obj, testName)
    If Not (obj Is Nothing) Then
        testsPassed = testsPassed + 1
        testResults = testResults & "<div style='color: green;'>[PASS] " & testName & "</div>"
    Else
        testsFailed = testsFailed + 1
        testResults = testResults & "<div style='color: red;'>[FAIL] " & testName & " (expected non-null object)</div>"
    End If
End Function

%>
<!DOCTYPE html>
<html>
<head>
    <title>AxonASP WScript.Shell Testing</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        h1 { color: #333; }
        .test-section { background: white; padding: 20px; margin: 20px 0; border-left: 4px solid #007bff; }
        .results { margin-top: 20px; padding: 15px; background: white; border-radius: 4px; }
        .summary { font-size: 18px; font-weight: bold; margin: 20px 0; }
        .pass { color: green; }
        .fail { color: red; }
        pre { background: #f8f9fa; padding: 15px; border-radius: 4px; overflow-x: auto; }
    </style>
</head>
<body>
    <h1>AxonASP WScript.Shell & AxExecute Testing</h1>
    
    <%
    ' Test 1: CreateObject - WScript.Shell
    %>
    <div class="test-section">
        <h2>Test 1: Object Creation</h2>
        <%
        On Error Resume Next
        
        Set objShell = Server.CreateObject("WScript.Shell")
        AssertNotNull objShell, "CreateObject('WScript.Shell')"
        
        Set objShell2 = Server.CreateObject("Wscript.Shell")
        AssertNotNull objShell2, "CreateObject('Wscript.Shell') - case insensitive"
        
        Set objShell3 = Server.CreateObject("Shell")
        AssertNotNull objShell3, "CreateObject('Shell') - shorthand"
        
        On Error GoTo 0
        %>
    </div>
    
    <%
    ' Test 2: Basic Run - Async (fire and forget)
    %>
    <div class="test-section">
        <h2>Test 2: Run Method (Basic)</h2>
        <%
        On Error Resume Next
        
        Set objShell = Server.CreateObject("WScript.Shell")
        intResult = objShell.Run("echo test", 0, True)
        
        AssertEqual intResult, 0, "Run with WaitOnReturn=True should return exit code 0"
        
        On Error GoTo 0
        %>
    </div>
    
    <%
    ' Test 3: Exec Method - Read output
    %>
    <div class="test-section">
        <h2>Test 3: Exec Method - StdOut.ReadAll()</h2>
        <%
        On Error Resume Next
        
        Set objShell = Server.CreateObject("WScript.Shell")
        
        ' Test with simple echo-like command
        Set objExec = objShell.Exec("echo Hello World")
        AssertNotNull objExec, "Exec should return WScriptExecObject"
        
        If Not (objExec Is Nothing) Then
            ' Read output
            strOutput = objExec.StdOut.ReadAll()
            AssertTrue (Len(strOutput) > 0), "StdOut.ReadAll() should return non-empty string"
            
            Response.Write "<p>Output from 'echo Hello World':</p>"
            Response.Write "<pre>" & strOutput & "</pre>"
        End If
        
        On Error GoTo 0
        %>
    </div>
    
    <%
    ' Test 4: Exec Method - AtEndOfStream property
    %>
    <div class="test-section">
        <h2>Test 4: Exec Method - AtEndOfStream Property</h2>
        <%
        On Error Resume Next
        
        Set objShell = Server.CreateObject("WScript.Shell")
        Set objExec = objShell.Exec("echo Line1" & vbCrLf & "echo Line2")
        
        AssertNotNull objExec, "Exec should return object"
        
        If Not (objExec Is Nothing) Then
            Set objStdOut = objExec.StdOut
            AssertTrue (Not objStdOut.AtEndOfStream), "Initially, AtEndOfStream should be False"
            
            ' Read content to reach end
            lineCount = 0
            Do While Not objStdOut.AtEndOfStream
                strLine = objStdOut.ReadLine()
                lineCount = lineCount + 1
                If lineCount > 100 Then ' Safety break
                    Exit Do
                End If
            Loop
            
            Response.Write "<p>Read " & lineCount & " lines from output</p>"
        End If
        
        On Error GoTo 0
        %>
    </div>
    
    <%
    ' Test 5: Exec with Status property
    %>
    <div class="test-section">
        <h2>Test 5: WScriptExecObject.Status Property</h2>
        <%
        On Error Resume Next
        
        Set objShell = Server.CreateObject("WScript.Shell")
        Set objExec = objShell.Exec("echo test")
        
        If objExec Is Not Nothing Then
            ' Process may be running or done depending on timing
            intStatus = objExec.Status
            AssertTrue ((intStatus = 0 Or intStatus = 1)), "Status should be 0 (running) or 1 (done)"
            
            ' Wait and check again
            objExec.WaitUntilDone()
            intStatus = objExec.Status
            AssertEqual intStatus, 1, "Status should be 1 (done) after WaitUntilDone()"
        End If
        
        On Error GoTo 0
        %>
    </div>
    
    <%
    ' Test 6: Exec with ExitCode property
    %>
    <div class="test-section">
        <h2>Test 6: WScriptExecObject.ExitCode Property</h2>
        <%
        On Error Resume Next
        
        Set objShell = Server.CreateObject("WScript.Shell")
        Set objExec = objShell.Exec("echo test")
        
        If Not (objExec Is Nothing) Then
            objExec.WaitUntilDone()
            intExitCode = objExec.ExitCode
            
            ' On most systems, a successful echo returns 0
            AssertTrue ((intExitCode >= 0)), "ExitCode should be non-negative (" & intExitCode & ")"
            
            Response.Write "<p>Exit Code: " & intExitCode & "</p>"
        End If
        
        On Error GoTo 0
        %>
    </div>
    
    <%
    ' Test 7: Exec with ProcessID property
    %>
    <div class="test-section">
        <h2>Test 7: WScriptExecObject.ProcessID Property</h2>
        <%
        On Error Resume Next
        
        Set objShell = Server.CreateObject("WScript.Shell")
        Set objExec = objShell.Exec("echo test")
        
        If Not (objExec Is Nothing) Then
            intPID = objExec.ProcessID
            AssertTrue ((intPID > 0)), "ProcessID should be positive (" & intPID & ")"
            Response.Write "<p>Process ID: " & intPID & "</p>"
        End If
        
        On Error GoTo 0
        %>
    </div>
    
    <%
    ' Test 8: AxExecute function
    %>
    <div class="test-section">
        <h2>Test 8: AxExecute Function (Basic)</h2>
        <%
        On Error Resume Next
        
        strOutput = AxExecute("echo Hello from AxExecute")
        
        AssertTrue (strOutput <> False And Len(strOutput) > 0), "AxExecute should return command output"
        
        Response.Write "<p>AxExecute output:</p>"
        Response.Write "<pre>" & Err.Number & ": " & strOutput & "</pre>"
        
        On Error GoTo 0
        %>
    </div>
    
    <%
    ' Test 9: GetEnv method
    %>
    <div class="test-section">
        <h2>Test 9: GetEnv Method</h2>
        <%
        On Error Resume Next
        
        Set objShell =  Server.CreateObject("WScript.Shell")
        
        ' Try to get PATH variable
        strPath = objShell.GetEnv("PATH")
        AssertTrue (Len(strPath) > 0), "GetEnv('PATH') should return non-empty value"
        
        ' Try to get a variable that typically exists
        strTemp = objShell.GetEnv("TEMP")
        Response.Write "<p>TEMP directory: " & strTemp & "</p>"
        
        ' Try non-existent variable
        strNonExist = objShell.GetEnv("THISVARIABLEDOESNOTEXIST12345")
        AssertTrue (Len(strNonExist) = 0), "GetEnv for non-existent variable should return empty string"
        
        On Error GoTo 0
        %>
    </div>
    
    <%
    ' Test 10: Multiple Exec instances
    %>
    <div class="test-section">
        <h2>Test 10: Multiple Concurrent Exec Instances</h2>
        <%
        On Error Resume Next
        
        Set objShell = CreateObject("WScript.Shell")
        Set objExec1 = objShell.Exec("echo First")
        Set objExec2 = objShell.Exec("echo Second")
        Set objExec3 = objShell.Exec("echo Third")
        
        AssertNotNull objExec1, "First Exec should succeed"
        AssertNotNull objExec2, "Second Exec should succeed"
        AssertNotNull objExec3, "Third Exec should succeed"
        
        If Not (objExec1 Is Nothing) And Not (objExec2 Is Nothing) And Not (objExec3 Is Nothing) Then
            objExec1.WaitUntilDone()
            objExec2.WaitUntilDone()
            objExec3.WaitUntilDone()
            
            str1 = objExec1.StdOut.ReadAll()
            str2 = objExec2.StdOut.ReadAll()
            str3 = objExec3.StdOut.ReadAll()
            
            AssertTrue ((Len(str1) > 0) And (Len(str2) > 0) And (Len(str3) > 0)), "All concurrent executions should produce output"
        End If
        
        On Error GoTo 0
        %>
    </div>
    
    <%
    ' Test Summary
    %>
    <div class="results">
        <div class="summary">
            Test Results: <span class="pass"><%= testsPassed %> Passed</span> | <span class="fail"><%= testsFailed %> Failed</span>
        </div>
        <% If testsFailed = 0 Then %>
            <p style="color: green; font-weight: bold;">All tests passed!</p>
        <% Else %>
            <p style="color: red; font-weight: bold;"><%= testsFailed %> test(s) failed</p>
        <% End If %>
        
        <h3>Detailed Results:</h3>
        <%= testResults %>
    </div>
    
    <hr>
    
    <%
    ' Additional Documentation
    %>
    <div class="test-section">
        <h2>WScript.Shell Documentation</h2>
        <p>For complete documentation, see: <a href="../../docs/WSCRIPT_SHELL_IMPLEMENTATION.md">WSCRIPT_SHELL_IMPLEMENTATION.md</a></p>
        
        <h3>Quick Reference:</h3>
        <ul>
            <li><strong>Run()</strong> - Execute command synchronously or asynchronously</li>
            <li><strong>Exec()</strong> - Execute command with stream access (stdout, stderr, stdin)</li>
            <li><strong>GetEnv()</strong> - Get environment variable value</li>
            <li><strong>WScriptExecObject.Status</strong> - 0=running, 1=done</li>
            <li><strong>WScriptExecObject.ExitCode</strong> - Exit code of the process</li>
            <li><strong>WScriptExecObject.ProcessID</strong> - Process ID</li>
            <li><strong>WScriptExecObject.StdOut</strong> - TextStream for reading output</li>
            <li><strong>WScriptExecObject.StdErr</strong> - TextStream for reading errors</li>
            <li><strong>WScriptExecObject.StdIn</strong> - TextStream for writing input</li>
            <li><strong>TextStream.ReadAll()</strong> - Read all remaining content</li>
            <li><strong>TextStream.ReadLine()</strong> - Read one line</li>
            <li><strong>TextStream.AtEndOfStream</strong> - Check if at end of stream</li>
        </ul>
    </div>
    
    <div class="test-section">
        <h2>AxExecute Function Documentation</h2>
        <p>For custom ASP/AxonASP functions, see: <a href="../../docs/CUSTOM_FUNCTIONS.md#axexecute">CUSTOM_FUNCTIONS.md#axexecute</a></p>
        
        <h3>Usage:</h3>
        <pre>strOutput = AxExecute(strCommand, [arrOutput], [intResultCode])</pre>
        
        <h3>Key Features:</h3>
        <ul>
            <li>Simple command execution with output capture</li>
            <li>Similar to PHP's exec() function</li>
            <li>Cross-platform (Windows, Unix, Linux, macOS)</li>
            <li>Returns command output as string</li>
            <li>Returns False on error</li>
        </ul>
    </div>

</body>
</html>
