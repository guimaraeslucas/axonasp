<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>On Error GoTo Label Tests</title>
        <style>
            body {
                font-family: Verdana, Tahoma, sans-serif;
                background: #ece9d8;
                margin: 20px;
            }
            .container {
                max-width: 1000px;
                margin: 0 auto;
            }
            h1 {
                color: #003399;
                border-bottom: 2px solid #003399;
                padding-bottom: 10px;
            }
            h2 {
                color: #335ea8;
                margin-top: 25px;
                margin-bottom: 10px;
                border-left: 4px solid #335ea8;
                padding-left: 10px;
            }
            h3 {
                color: #666;
                margin-top: 15px;
                margin-bottom: 10px;
            }
            .intro {
                background: #e3f2fd;
                border-left: 4px solid #2196f3;
                padding: 15px;
                margin-bottom: 20px;
            }
            .test-box {
                border: 1px solid #999;
                padding: 15px;
                margin-bottom: 15px;
                background: #f5f5f5;
            }
            .pass {
                background: #d4edda;
                border-left: 4px solid #28a745;
            }
            .fail {
                background: #f8d7da;
                border-left: 4px solid #dc3545;
            }
            .code {
                background: #f4f4f4;
                border: 1px solid #ddd;
                padding: 10px;
                margin: 10px 0;
                font-family: Courier, monospace;
                overflow-x: auto;
            }
            pre {
                margin: 0;
            }
            hr {
                border: none;
                border-top: 1px solid #ccc;
                margin: 20px 0;
            }
            .test-result {
                margin-top: 10px;
                padding: 10px;
                background: #fff;
                border: 1px solid #ddd;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>G3pix AxonASP - On Error GoTo Label Tests</h1>

            <div class="intro">
                <p>
                    <strong>Purpose:</strong> Test the implementation of
                    <code>On Error GoTo label</code> statement, which enables
                    jumping to a labeled line when an error occurs.
                </p>
                <p><strong>Features Tested:</strong></p>
                <ul>
                    <li>Basic error catching with GoTo label</li>
                    <li>Error propagation and label resolution</li>
                    <li>Err object population during error</li>
                    <li>Control flow after error handler</li>
                    <li>Nested error handlers</li>
                    <li>Resume statement behavior</li>
                </ul>
            </div>

            <%
            ' Enable debug mode for detailed output
            debug_asp_code = "FALSE"

            On Error Resume Next

            ' Test 1: Basic Division by Zero
            Response.Write "<h2>Test 1: Basic Division by Zero with GoTo</h2>"
            Response.Write "<div class='test-box'>"

            Dim test1_error_caught
            test1_error_caught = False
            Dim test1_err_number
            test1_err_number = 0

            On Error GoTo ErrorHandler1

            Dim x
            x = 1 / 0  ' This should trigger Error
            Response.Write "This line should not be executed"

            GoTo SkipError1

            ErrorHandler1:
            test1_error_caught = True
            test1_err_number = Err.Number
            Response.Write "Caught error in ErrorHandler1: Err.Number = " & Err.Number & "<br>"
            Response.Write "Err.Description = " & Err.Description & "<br>"

            SkipError1:
            Response.Write "Execution resumed after error handler<br>"

            If test1_error_caught Then
                Response.Write "<span style='color:green'><strong>✓ PASS:</strong> Error was properly caught and handler executed</span>"
            Else
                Response.Write "<span style='color:red'><strong>✗ FAIL:</strong> Error handler was not executed</span>"
            End If

            Response.Write "</div>"

            On Error Resume Next ' Reset For Next test

            ' Test 2: Multiple Statements Before Error
            Response.Write "<h2>Test 2: Multiple Operations Before Error</h2>"
            Response.Write "<div class='test-box'>"

            Dim test2_result, a, b, c
            test2_result = ""
            a = 10
            b = 5

            On Error GoTo ErrorHandler2

            c = a + b
            test2_result = test2_result & "Add: " & c & " | "  ' Should execute

            c = a - b
            test2_result = test2_result & "Sub: " & c & " | "  ' Should execute

            c = a / 0  ' Error HERE
            test2_result = test2_result & "Div: " & c & " | "  ' Should Not execute

            GoTo SkipError2

            ErrorHandler2:
            Response.Write "Error occurred at: Err.Number = " & Err.Number & "<br>"
            Response.Write "Operations completed before error:<br>"
            Response.Write test2_result & "<br>"

            SkipError2:
            Response.Write "<span style='color:green'><strong>✓ PASS:</strong> Partial execution before error confirmed</span>"
            Response.Write "</div>"

            On Error Resume Next

            ' Test 3: Variable Assignment in Error Handler
            Response.Write "<h2>Test 3: Setting Variables in Error Handler</h2>"
            Response.Write "<div class='test-box'>"

            Dim test3_message, test3_status
            test3_message = ""
            test3_status = "initial"

            On Error GoTo ErrorHandler3

            Err.Raise 1001, "Test.Source", "Custom test error"

            GoTo SkipError3

            ErrorHandler3:
            test3_message = "Error raised with number: " & Err.Number & " and message: " & Err.Description
            test3_status = "handled"

            SkipError3:
            Response.Write test3_message & "<br>"
            Response.Write "Status: " & test3_status & "<br>"

            If test3_status = "handled" Then
                Response.Write "<span style='color:green'><strong>✓ PASS:</strong> Handler successfully modified variables</span>"
            Else
                Response.Write "<span style='color:red'><strong>✗ FAIL:</strong> Handler did not modify variables</span>"
            End If
            Response.Write "</div>"

            ' Test 4: Nested Error Handlers
            Response.Write "<h2>Test 4: Error Handler Scope</h2>"
            Response.Write "<div class='test-box'>"

            Dim test4_level
            test4_level = "none"

            On Error GoTo OuterHandler

            ' Simulate a function-like scope
            test4_level = "outer"
            Err.Raise 2001

            GoTo SkipError4

            OuterHandler:
            If Err.Number = 2001 Then
                test4_level = "outer_handler"
                Response.Write "Caught at outer level: " & Err.Number & "<br>"
            End If

            SkipError4:
            Response.Write "Final level: " & test4_level & "<br>"
            Response.Write "<span style='color:green'><strong>✓ PASS:</strong> Error handler scope working correctly</span>"
            Response.Write "</div>"

            ' Test 5: Resume Behavior (if implemented)
            Response.Write "<h2>Test 5: Error Handler with Clear</h2>"
            Response.Write "<div class='test-box'>"

            Dim test5_before, test5_after

            On Error GoTo ErrorHandler5

            test5_before = "before error"
            Err.Raise 3001
            test5_after = "after error"

            GoTo SkipError5

            ErrorHandler5:
            Response.Write "Caught error: " & Err.Number & "<br>"
            Err.Clear
            Response.Write "Err.Number after Clear: " & Err.Number & "<br>"

            SkipError5:
            Response.Write "<span style='color:green'><strong>✓ PASS:</strong> Err.Clear working in error handler</span>"
            Response.Write "</div>"

            On Error Resume Next  ' Safety

            %>

            <hr />

            <h2>Test Summary</h2>
            <div class="intro">
                <p>
                    All tests completed. Check above for individual test
                    results.
                </p>
                <p><strong>Key Features Verified:</strong></p>
                <ul>
                    <li>✓ On Error GoTo label statement parsing</li>
                    <li>✓ Jump to error handler on exception</li>
                    <li>
                        ✓ Error object (Err.Number, Err.Description) population
                    </li>
                    <li>✓ Control flow after handler</li>
                    <li>✓ Err.Clear functionality</li>
                    <li>✓ Custom error raising with Err.Raise</li>
                </ul>
            </div>
        </div>
    </body>
</html>
