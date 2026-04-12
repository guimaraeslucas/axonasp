<%
Option Explicit
Response.LCID = 1046
Response.buffer = True
%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <title>Operator Spacing Test</title>
        <style type="text/css">
            body {
                font-family: Arial, Helvetica, sans-serif;
            }
            .pass {
                color: green;
            }
            .fail {
                color: red;
            }
            pre {
                border: 1px solid #ccc;
                background: #f0f0f0;
                padding: 10px;
            }
        </style>
    </head>
    <body>
        <h1>Operator Spacing Test</h1>
        <p>Testing compound operators with whitespace between components</p>

        <%
        Dim testsPassed, testsFailed
        testsPassed = 0
        testsFailed = 0

        ' Test 1: >= with space
        Dim result1, expected1
        expected1 = True
        result1 = False
        If 5 >  = 3 Then result1 = True

        If result1 = expected1 Then
            Response.Write "<p class='pass'>✓ Test 1 PASS: 5 > = 3 evaluates to true</p>"
            testsPassed = testsPassed + 1
        Else
            Response.Write "<p class='fail'>✗ Test 1 FAIL: 5 > = 3 should be true but got false</p>"
            testsFailed = testsFailed + 1
        End If

        ' Test 2: <= with space
        Dim result2, expected2
        expected2 = True
        result2 = False
        If 3 <  = 5 Then result2 = True

        If result2 = expected2 Then
            Response.Write "<p class='pass'>✓ Test 2 PASS: 3 < = 5 evaluates to true</p>"
            testsPassed = testsPassed + 1
        Else
            Response.Write "<p class='fail'>✗ Test 2 FAIL: 3 < = 5 should be true but got false</p>"
            testsFailed = testsFailed + 1
        End If

        ' Test 3: >= with multiple spaces
        Dim result3, expected3
        expected3 = True
        result3 = False
        If 10 >  = 10 Then result3 = True

        If result3 = expected3 Then
            Response.Write "<p class='pass'>✓ Test 3 PASS: 10 >  = 10 evaluates to true</p>"
            testsPassed = testsPassed + 1
        Else
            Response.Write "<p class='fail'>✗ Test 3 FAIL: 10 >  = 10 should be true but got false</p>"
            testsFailed = testsFailed + 1
        End If

        ' Test 4: <= with multiple spaces
        Dim result4, expected4
        expected4 = True
        result4 = False
        If 2 <  = 2 Then result4 = True

        If result4 = expected4 Then
            Response.Write "<p class='pass'>✓ Test 4 PASS: 2 <  = 2 evaluates to true</p>"
            testsPassed = testsPassed + 1
        Else
            Response.Write "<p class='fail'>✗ Test 4 FAIL: 2 <  = 2 should be true but got false</p>"
            testsFailed = testsFailed + 1
        End If

        ' Test 5: >= without space (normal syntax)
        Dim result5, expected5
        expected5 = True
        result5 = False
        If 7 >  = 4 Then result5 = True

        If result5 = expected5 Then
            Response.Write "<p class='pass'>✓ Test 5 PASS: 7 >= 4 evaluates to true</p>"
            testsPassed = testsPassed + 1
        Else
            Response.Write "<p class='fail'>✗ Test 5 FAIL: 7 >= 4 should be true but got false</p>"
            testsFailed = testsFailed + 1
        End If

        ' Test 6: <= without space (normal syntax)
        Dim result6, expected6
        expected6 = True
        result6 = False
        If 1 <  = 9 Then result6 = True

        If result6 = expected6 Then
            Response.Write "<p class='pass'>✓ Test 6 PASS: 1 <= 9 evaluates to true</p>"
            testsPassed = testsPassed + 1
        Else
            Response.Write "<p class='fail'>✗ Test 6 FAIL: 1 <= 9 should be true but got false</p>"
            testsFailed = testsFailed + 1
        End If

        ' Test 7: > with space > (not >=)
        Dim result7, expected7
        expected7 = True
        result7 = False
        If 5 > 3 Then result7 = True

        If result7 = expected7 Then
            Response.Write "<p class='pass'>✓ Test 7 PASS: 5 > 3 evaluates to true</p>"
            testsPassed = testsPassed + 1
        Else
            Response.Write "<p class='fail'>✗ Test 7 FAIL: 5 > 3 should be true but got false</p>"
            testsFailed = testsFailed + 1
        End If

        ' Summary
        Response.Write "<hr>"
        Response.Write "<h2>Test Results</h2>"
        Response.Write "<p>Passed: <span class='pass'>" & testsPassed & "</span></p>"
        Response.Write "<p>Failed: <span class='fail'>" & testsFailed & "</span></p>"
        If testsFailed = 0 Then
            Response.Write "<p class='pass'><strong>All tests passed!</strong></p>"
        Else
            Response.Write "<p class='fail'><strong>Some tests failed!</strong></p>"
        End If
        %>
    </body>
</html>
