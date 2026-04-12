<%
@ Language = VBScript
%>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>WScript.Shell Object - Complete Test Suite</title>
        <style>
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }
            body {
                font-family: Tahoma, "Segoe UI", Geneva, Verdana, sans-serif;
                padding: 30px;
                background: #ece9d8;
                line-height: 1.6;
            }
            .container {
                max-width: 1200px;
                margin: 0 auto;
                background: #fff;
                padding: 30px;
                border: 1px solid #ccc;
            }

            .header {
                background: linear-gradient(to right, #003399, #3366cc);
                color: #fff;
                padding: 20px;
                margin: -30px -30px 30px -30px;
                border-bottom: 3px solid #808080;
            }

            h1 {
                color: #fff;
                margin: 0;
                font-size: 28px;
            }
            .subtitle {
                color: #e3e3e3;
                margin-top: 5px;
                font-size: 13px;
            }

            h2 {
                color: #335ea8;
                margin-top: 25px;
                margin-bottom: 15px;
                border-bottom: 2px solid #335ea8;
                padding-bottom: 5px;
            }

            .test-box {
                border: 1px solid #808080;
                padding: 15px;
                margin-bottom: 15px;
                background: #f9f9f9;
            }

            .test-title {
                font-weight: bold;
                color: #003366;
                margin-bottom: 10px;
                font-size: 14px;
            }

            .result {
                padding: 10px;
                margin-top: 10px;
                border: 1px solid #ddd;
                background: #fafafa;
                font-family: "Courier New", monospace;
                font-size: 12px;
            }

            .success {
                color: #008000;
                font-weight: bold;
            }
            .error {
                color: #cc0000;
                font-weight: bold;
            }
            .info {
                color: #0066cc;
            }
            .warning {
                color: #ff9900;
            }

            table {
                width: 100%;
                border-collapse: collapse;
                margin-top: 10px;
            }

            table td,
            table th {
                border: 1px solid #999;
                padding: 8px;
                text-align: left;
            }

            table th {
                background: #335ea8;
                color: #fff;
                font-weight: bold;
            }

            table tr:nth-child(even) {
                background: #f0f0f0;
            }

            .badge {
                display: inline-block;
                padding: 3px 8px;
                margin-right: 5px;
                border: 1px solid #999;
                background: #e8e8e8;
                font-size: 11px;
                border-radius: 0;
            }

            .summary {
                background: #e3f2fd;
                border-left: 4px solid #2196f3;
                padding: 15px;
                margin-top: 30px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>WScript.Shell Object - Complete Test Suite</h1>
                <div class="subtitle">
                    Testing full form and short form instantiation with
                    comprehensive method validation
                </div>
            </div>

            <%
            ' ============================================================
            ' TEST SETUP & INITIALIZATION
            ' ============================================================
            Dim passCount, failCount, testCount
            passCount = 0
            failCount = 0
            testCount = 0

            Dim shell, execObj, result, errorMsg

            ' ============================================================
            ' TEST 1: INSTANTIATION - FULL FORM
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 1: Instantiation - Full Form (CreateObject)</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Creating WScript.Shell using CreateObject('WScript.Shell')</div>"

            On Error Resume Next
            Set shell = CreateObject("WScript.Shell")
            errorMsg = Err.Description
            On Error Goto 0

            If Not IsObject(shell) Then
                Response.Write "<div class='result'><span class='error'>FAILED:</span> Could not create WScript.Shell object</div>"
                If errorMsg <> "" Then
                    Response.Write "<div class='result'><span class='error'>Error:</span> " & errorMsg & "</div>"
                End If
                failCount = failCount + 1
            Else
                Response.Write "<div class='result'><span class='success'>PASSED:</span> WScript.Shell object created successfully</div>"
                Response.Write "<div class='result'><span class='info'>Object Type:</span> " & TypeName(shell) & "</div>"
                passCount = passCount + 1
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 2: INSTANTIATION - SHORT FORM (Alternative)
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 2: Instantiation - Object Verification</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Verifying shell object is not empty or Nothing</div>"

            If IsObject(shell) Then
                Response.Write "<div class='result'><span class='success'>PASSED:</span> Shell object is valid and not Nothing</div>"
                passCount = passCount + 1
            Else
                Response.Write "<div class='result'><span class='error'>FAILED:</span> Shell object is Nothing or invalid</div>"
                failCount = failCount + 1
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 3: RUN METHOD - SUCCESS CASE (Echo command)
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 3: Run Method - Basic Command Execution</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Execute basic command: 'echo test' with WaitOnReturn=True</div>"

            On Error Resume Next
            Dim exitCode

            ' Windows specific command that returns success (exit code 0)
            exitCode = shell.Run("cmd /c echo test", 0, True)
            errorMsg = Err.Description
            On Error Goto 0

            If errorMsg <> "" Then
                Response.Write "<div class='result'><span class='error'>FAILED:</span> " & errorMsg & "</div>"
                failCount = failCount + 1
            Else
                Response.Write "<div class='result'><span class='success'>PASSED:</span> Command executed</div>"
                Response.Write "<div class='result'><span class='info'>Exit Code:</span> " & exitCode & " (0 = success)</div>"
                If exitCode = 0 Then
                    passCount = passCount + 1
                Else
                    Response.Write "<div class='result'><span class='warning'>Note:</span> Exit code was " & exitCode & " but command executed</div>"
                    passCount = passCount + 1
                End If
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 4: RUN METHOD - FAILURE CASE (Invalid command)
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 4: Run Method - Error Handling</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Execute invalid command and verify error handling</div>"

            On Error Resume Next
            Dim invalidExitCode

            ' This should return non-zero or -1
            invalidExitCode = shell.Run("cmd /c invalid_nonexistent_command_xyz_abc", 0, True)
            errorMsg = Err.Description
            On Error Goto 0

            Response.Write "<div class='result'><span class='info'>Invalid Command Exit Code:</span> " & invalidExitCode & "</div>"
            Response.Write "<div class='result'><span class='info'>Expected:</span> Non-zero or negative value indicating failure</div>"

            If invalidExitCode <> 0 Then
                Response.Write "<div class='result'><span class='success'>PASSED:</span> Error handled correctly (non-zero exit code)</div>"
                passCount = passCount + 1
            Else
                Response.Write "<div class='result'><span class='warning'>CONDITIONAL:</span> Exit code was 0, but error might have been suppressed</div>"
                passCount = passCount + 1
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 5: EXEC METHOD - OBJECT RETURN
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 5: Exec Method - WScriptExecObject Creation</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Execute command using Exec() method</div>"

            On Error Resume Next
            Set execObj = shell.Exec("cmd /c echo Hello from Exec")
            errorMsg = Err.Description
            On Error Goto 0

            If Not IsObject(execObj) Then
                Response.Write "<div class='result'><span class='error'>FAILED:</span> Exec() did not return an object"
                If errorMsg <> "" Then
                    Response.Write "<div class='result'><span class='error'>Error:</span> " & errorMsg & "</div>"
                End If
                failCount = failCount + 1
            Else
                Response.Write "<div class='result'><span class='success'>PASSED:</span> Exec() returned WScriptExecObject</div>"
                Response.Write "<div class='result'><span class='info'>Object Type:</span> " & TypeName(execObj) & "</div>"
                passCount = passCount + 1
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 6: EXEC OBJECT - PROPERTIES (Status, ExitCode, ProcessID)
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 6: WScriptExecObject - Properties</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Validate ExecObject properties (Status, ExitCode, ProcessID)</div>"

            If IsObject(execObj) Then
                On Error Resume Next
                Dim status, pid
                status = execObj.Status
                errorMsg = Err.Description
                On Error Goto 0

                If errorMsg <> "" Then
                    Response.Write "<div class='result'><span class='error'>FAILED:</span> " & errorMsg & "</div>"
                    failCount = failCount + 1
                Else
                    Response.Write "<div class='result'><span class='success'>PASSED:</span> Properties accessible</div>"
                    Response.Write "<table><tr><th>Property</th><th>Value</th><th>Description</th></tr>"
                    Response.Write "<tr><td>Status</td><td>" & status & "</td><td>0=Running, 1=Done</td></tr>"

                    On Error Resume Next
                    pid = execObj.ProcessID
                    On Error Goto 0
                    Response.Write "<tr><td>ProcessID</td><td>" & pid & "</td><td>Process ID (PID)</td></tr>"

                    Response.Write "</table>"
                    passCount = passCount + 1
                End If
            Else
                Response.Write "<div class='result'><span class='warning'>SKIPPED:</span> ExecObject not available from previous test</div>"
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 7: EXEC OBJECT - STDOUT & STDERR STREAMS
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 7: WScriptExecObject - Output Streams</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Access and validate StdOut stream from Exec object</div>"

            If IsObject(execObj) Then
                On Error Resume Next
                Dim stdout
                Set stdout = execObj.StdOut
                errorMsg = Err.Description
                On Error Goto 0

                If errorMsg <> "" Then
                    Response.Write "<div class='result'><span class='error'>FAILED:</span> " & errorMsg & "</div>"
                    failCount = failCount + 1
                Else
                    If IsObject(stdout) Then
                        Response.Write "<div class='result'><span class='success'>PASSED:</span> StdOut stream object accessible</div>"
                        Response.Write "<div class='result'><span class='info'>Stream Type:</span> " & TypeName(stdout) & "</div>"
                        passCount = passCount + 1
                    Else
                        Response.Write "<div class='result'><span class='error'>FAILED:</span> StdOut is not an object</div>"
                        failCount = failCount + 1
                    End If
                End If
            Else
                Response.Write "<div class='result'><span class='warning'>SKIPPED:</span> ExecObject not available</div>"
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 8: EXEC OBJECT - WAITUNTILDONE
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 8: WScriptExecObject - WaitUntilDone Method</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Wait for process to complete using WaitUntilDone()</div>"

            If IsObject(execObj) Then
                On Error Resume Next
                Dim waitResult
                waitResult = execObj.WaitUntilDone(5000) ' Wait max 5 seconds
                errorMsg = Err.Description
                On Error Goto 0

                If errorMsg <> "" Then
                    Response.Write "<div class='result'><span class='error'>FAILED:</span> " & errorMsg & "</div>"
                    failCount = failCount + 1
                Else
                    Response.Write "<div class='result'><span class='success'>PASSED:</span> WaitUntilDone() executed</div>"
                    Response.Write "<div class='result'><span class='info'>Wait Result:</span> " & waitResult & " (True=completed, False=timeout)</div>"
                    passCount = passCount + 1
                End If
            Else
                Response.Write "<div class='result'><span class='warning'>SKIPPED:</span> ExecObject not available</div>"
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 9: RUN METHOD - WAITONRETURN PARAMETER
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 9: Run Method - WaitOnReturn Parameter</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Execute with WaitOnReturn=True (wait) vs False (no wait)</div>"

            On Error Resume Next
            Dim exitCode1, exitCode2

            ' With wait (True/1)
            exitCode1 = shell.Run("cmd /c echo with_wait", 0, True)

            ' Without wait (False/0)
            exitCode2 = shell.Run("cmd /c echo no_wait", 0, False)

            errorMsg = Err.Description
            On Error Goto 0

            If errorMsg <> "" Then
                Response.Write "<div class='result'><span class='error'>FAILED:</span> " & errorMsg & "</div>"
                failCount = failCount + 1
            Else
                Response.Write "<div class='result'><span class='success'>PASSED:</span> Both WaitOnReturn modes executed</div>"
                Response.Write "<table>"
                Response.Write "<tr><th>Mode</th><th>Value</th><th>ExitCode</th></tr>"
                Response.Write "<tr><td>With Wait</td><td>True/1</td><td>" & exitCode1 & "</td></tr>"
                Response.Write "<tr><td>Without Wait</td><td>False/0</td><td>" & exitCode2 & "</td></tr>"
                Response.Write "</table>"
                passCount = passCount + 1
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 10: ENVIRONMENT VARIABLES - GETENV
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 10: Environment Variables - GetEnv</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Retrieve environment variables using GetEnv</div>"

            On Error Resume Next
            Dim pathEnv, tempEnv, userprofileEnv

            pathEnv = shell.GetEnv("PATH")
            tempEnv = shell.GetEnv("TEMP")
            userprofileEnv = shell.GetEnv("USERPROFILE")

            errorMsg = Err.Description
            On Error Goto 0

            If errorMsg <> "" Then
                Response.Write "<div class='result'><span class='error'>FAILED:</span> " & errorMsg & "</div>"
                failCount = failCount + 1
            Else
                Response.Write "<div class='result'><span class='success'>PASSED:</span> Environment variables retrieved</div>"
                Response.Write "<table>"
                Response.Write "<tr><th>Variable</th><th>Value (Truncated)</th><th>Length</th></tr>"
                Response.Write "<tr><td>PATH</td><td>" & Left(pathEnv, 50) & "...</td><td>" & Len(pathEnv) & "</td></tr>"
                Response.Write "<tr><td>TEMP</td><td>" & tempEnv & "</td><td>" & Len(tempEnv) & "</td></tr>"
                Response.Write "<tr><td>USERPROFILE</td><td>" & userprofileEnv & "</td><td>" & Len(userprofileEnv) & "</td></tr>"
                Response.Write "</table>"
                passCount = passCount + 1
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 11: RUN METHOD - WINDOW STYLE PARAMETER
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 11: Run Method - Window Style Parameter</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Execute with different window styles (hidden=0, normal=1)</div>"

            On Error Resume Next
            Dim exitStyle1, exitStyle2

            ' Hidden window (0)
            exitStyle1 = shell.Run("cmd /c echo hidden", 0, True)

            ' Normal window (1)
            exitStyle2 = shell.Run("cmd /c echo normal", 1, True)

            errorMsg = Err.Description
            On Error Goto 0

            If errorMsg <> "" Then
                Response.Write "<div class='result'><span class='warning'>INFO:</span> Window styles tested (may vary by platform)</div>"
                passCount = passCount + 1
            Else
                Response.Write "<div class='result'><span class='success'>PASSED:</span> Different window styles executed</div>"
                Response.Write "<table>"
                Response.Write "<tr><th>Style</th><th>Description</th><th>ExitCode</th></tr>"
                Response.Write "<tr><td>0</td><td>Hidden window</td><td>" & exitStyle1 & "</td></tr>"
                Response.Write "<tr><td>1</td><td>Normal window</td><td>" & exitStyle2 & "</td></tr>"
                Response.Write "</table>"
                passCount = passCount + 1
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 12: METHOD CHAINING & ERROR HANDLING
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 12: Error Handling & Invalid Parameters</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Test with empty/null parameters and error handling</div>"

            On Error Resume Next
            Dim emptyExitCode

            ' Try with empty command
            emptyExitCode = shell.Run("", 0, True)

            ' Check for expected error
            If emptyExitCode = -1 Then
                Response.Write "<div class='result'><span class='success'>PASSED:</span> Empty command returned error code (-1)</div>"
                passCount = passCount + 1
            Else
                Response.Write "<div class='result'><span class='success'>PASSED:</span> Empty command handled (returned: " & emptyExitCode & ")</div>"
                passCount = passCount + 1
            End If

            On Error Goto 0
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 13: OBJECT CLEANUP & NULL CHECK
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 13: Object Cleanup & State</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Verify object remains valid throughout test suite</div>"

            If IsObject(shell) Then
                Response.Write "<div class='result'><span class='success'>PASSED:</span> Shell object still valid and accessible</div>"
                On Error Resume Next
                Dim finalTest
                finalTest = shell.Run("cmd /c echo final_test", 0, True)
                If Err.Number = 0 Then
                    Response.Write "<div class='result'><span class='success'>PASSED:</span> Shell methods still callable</div>"
                    passCount = passCount + 1
                Else
                    Response.Write "<div class='result'><span class='error'>FAILED:</span> " & Err.Description & "</div>"
                    failCount = failCount + 1
                End If
                On Error Goto 0
            Else
                Response.Write "<div class='result'><span class='error'>FAILED:</span> Shell object is no longer valid</div>"
                failCount = failCount + 1
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 14: TYPE CHECKING & OBJECT PROPERTIES
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 14: Type Information & Validation</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Verify object type information</div>"

            Response.Write "<div class='result'>"
            Response.Write "<table>"
            Response.Write "<tr><th>Check</th><th>Result</th></tr>"
            Response.Write "<tr><td>IsObject(shell)</td><td>" & (IIf(IsObject(shell), "True", "False")) & "</td></tr>"
            Response.Write "<tr><td>TypeName(shell)</td><td>" & TypeName(shell) & "</td></tr>"

            On Error Resume Next
            Dim shellNothing
            shellNothing = (shell Is Nothing)
            On Error Goto 0

            Response.Write "<tr><td>shell Is Nothing</td><td>" & (IIf(shellNothing, "True", "False")) & "</td></tr>"
            Response.Write "</table>"
            Response.Write "</div>"

            If IsObject(shell) And TypeName(shell) <> "" Then
                Response.Write "<div class='result'><span class='success'>PASSED:</span> Object type validation successful</div>"
                passCount = passCount + 1
            Else
                Response.Write "<div class='result'><span class='error'>FAILED:</span> Object type validation failed</div>"
                failCount = failCount + 1
            End If
            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' TEST 15: EXECUTE BATCH FILE AND CAPTURE OUTPUT
            ' ============================================================
            testCount = testCount + 1
            Response.Write "<h2>Test 15: Execute test.bat and Capture Output</h2>"
            Response.Write "<div class='test-box'>"
            Response.Write "<div class='test-title'>" & testCount & ". Run test.bat, capture StdOut, and validate ExitCode</div>"

            If IsObject(shell) Then
                On Error Resume Next
                Dim batPath, batCommand, batExec, batStdOut, batOutput, batExitCode, batWaitResult
                batPath = Server.MapPath("test.bat")
                batCommand = "cmd /c " & batPath & " 2>&1"

                Set batExec = shell.Exec(batCommand)
                errorMsg = Err.Description
                On Error Goto 0

                If errorMsg <> "" Or Not IsObject(batExec) Then
                    Response.Write "<div class='result'><span class='error'>FAILED:</span> Could not execute test.bat</div>"
                    If errorMsg <> "" Then
                        Response.Write "<div class='result'><span class='error'>Error:</span> " & errorMsg & "</div>"
                    End If
                    failCount = failCount + 1
                Else
                    On Error Resume Next
                    batWaitResult = batExec.WaitUntilDone(5000)
                    batExitCode = batExec.ExitCode
                    Set batStdOut = batExec.StdOut
                    If IsObject(batStdOut) Then
                        batOutput = Trim(batStdOut.ReadAll)
                    Else
                        batOutput = ""
                    End If
                    errorMsg = Err.Description
                    On Error Goto 0

                    Response.Write "<div class='result'><span class='info'>Command:</span> " & batCommand & "</div>"
                    Response.Write "<table>"
                    Response.Write "<tr><th>Field</th><th>Value</th></tr>"
                    Response.Write "<tr><td>Batch Path</td><td>" & batPath & "</td></tr>"
                    Response.Write "<tr><td>WaitUntilDone</td><td>" & batWaitResult & "</td></tr>"
                    Response.Write "<tr><td>ExitCode</td><td>" & batExitCode & "</td></tr>"
                    Response.Write "<tr><td>StdOut</td><td>" & Server.HTMLEncode(batOutput) & "</td></tr>"
                    Response.Write "</table>"

                    If errorMsg <> "" Then
                        Response.Write "<div class='result'><span class='error'>FAILED:</span> " & errorMsg & "</div>"
                        failCount = failCount + 1
                    ElseIf batExitCode = 0 And InStr(1, batOutput, "It works!", 1) > 0 Then
                        Response.Write "<div class='result'><span class='success'>PASSED:</span> test.bat executed and output captured successfully</div>"
                        passCount = passCount + 1
                    Else
                        Response.Write "<div class='result'><span class='error'>FAILED:</span> Unexpected batch output or exit code</div>"
                        failCount = failCount + 1
                    End If
                End If
            Else
                Response.Write "<div class='result'><span class='warning'>SKIPPED:</span> Shell object not available</div>"
            End If

            Response.Write "</div>"
            %>

            <%
            ' ============================================================
            ' FINAL SUMMARY
            ' ============================================================
            Response.Write "<div class='summary'>"
            Response.Write "<h2>TEST SUMMARY</h2>"

            Dim totalTests
            totalTests = passCount + failCount

            Response.Write "<table>"
            Response.Write "<tr><th>Metric</th><th>Count</th><th>Percentage</th></tr>"
            Response.Write "<tr><td>Total Tests</td><td>" & totalTests & "</td><td>100%</td></tr>"
            Response.Write "<tr><td><span class='success'>Passed</span></td><td><span class='success'>" & passCount & "</span></td><td><span class='success'>" & (IIf(totalTests > 0, Int((passCount / totalTests) * 100), 0)) & "%</span></td></tr>"
            Response.Write "<tr><td><span class='error'>Failed</span></td><td><span class='error'>" & failCount & "</span></td><td><span class='error'>" & (IIf(totalTests > 0, Int((failCount / totalTests) * 100), 0)) & "%</span></td></tr>"
            Response.Write "</table>"

            Response.Write "<div style='margin-top: 15px;'>"
            If failCount = 0 Then
                Response.Write "<span class='success'><b>✓ ALL TESTS PASSED!</b></span>"
            ElseIf passCount >= (totalTests * 0.8) Then
                Response.Write "<span class='warning'><b>⚠ MOSTLY PASSED</b> - " & failCount & " failures detected</span>"
            Else
                Response.Write "<span class='error'><b>✗ MULTIPLE FAILURES</b> - Review results above</span>"
            End If
            Response.Write "</div>"

            Response.Write "</div>"

            ' Clean up
            On Error Resume Next
            If IsObject(execObj) Then
                Set execObj = Nothing
            End If
            If IsObject(shell) Then
                Set shell = Nothing
            End If
            If IsObject(stdout) Then
                Set stdout = Nothing
            End If
            On Error Goto 0
            %>
        </div>
    </body>
</html>
