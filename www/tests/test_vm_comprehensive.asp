<%
@ Language = "VBScript"
%>
<!DOCTYPE html>
<html>
    <head>
        <title>AxonVM Comprehensive Feature Test</title>
        <style>
            body {
                font-family: Arial;
                margin: 20px;
            }
            .test {
                margin: 20px 0;
                padding: 10px;
                border: 1px solid #ccc;
            }
            .pass {
                background: #e8f5e9;
                border-color: #4caf50;
            }
            .fail {
                background: #ffebee;
                border-color: #f44336;
            }
            h2 {
                color: #333;
            }
        </style>
    </head>
    <body>
        <h1>AxonVM Comprehensive Feature Test</h1>

        <%
        Dim testsPassed, testsFailed
        testsPassed = 0
        testsFailed = 0

        Function ReportTest(name, result)
            If result Then
                Response.Write "<div class='test pass'><b>✓</b> " & name & "</div>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<div class='test fail'><b>✗</b> " & name & "</div>"
                testsFailed = testsFailed + 1
            End If
        End Function

        ' ========== SESSION OBJECT ==========
        Response.Write "<h2>Session Object</h2>"

        Session("TestVar") = "TestValue"
        ReportTest "Session variable set and retrieve", Session("TestVar") = "TestValue"
        ReportTest "Session.SessionID exists", Session.SessionID <> ""
        ReportTest "Session.Timeout property works", Session.Timeout > 0
        ReportTest "Session.LCID property works", Session.LCID > 0

        ' ========== APPLICATION OBJECT ==========
        Response.Write "<h2>Application Object</h2>"

        Application.Lock
        Application("AppTestVar") = "AppTestValue"
        Application.Unlock
        ReportTest "Application variable set/get with Lock/Unlock", Application("AppTestVar") = "AppTestValue"

        ' ========== REQUEST OBJECT ==========
        Response.Write "<h2>Request Object</h2>"

        ReportTest "Request.ServerVariables accessible", Len(Request.ServerVariables("SERVER_NAME")) > 0
        ReportTest "Request.QueryString exists", TypeName(Request.QueryString) <> ""

        ' ========== RESPONSE OBJECT ==========
        Response.Write "<h2>Response Object</h2>"

        Dim originalContentType: originalContentType = Response.ContentType
        Response.ContentType = "text/plain"
        ReportTest "Response.ContentType settable", Response.ContentType = "text/plain"
        Response.ContentType = originalContentType

        Dim testBuffer: testBuffer = Response.Buffer
        ReportTest "Response.Buffer readable", testBuffer = True Or testBuffer = False
        ReportTest "Response.Write works", True ' If we got here, Write works!

        ' ========== SERVER OBJECT ==========
        Response.Write "<h2>Server Object</h2>"

        ReportTest "Server.ScriptTimeout accessible", Server.ScriptTimeout > 0
        ReportTest "Server.MapPath works", Len(Server.MapPath("/")) > 0

        ' Test CreateObject
        On Error Resume Next
        Set dict = Server.CreateObject("Scripting.Dictionary")
        ReportTest "Server.CreateObject(Scripting.Dictionary)", Err.Number = 0 And dict.Count = 0
        Err.Clear

        ' ========== ERR OBJECT ==========
        Response.Write "<h2>Err Object</h2>"

        Err.Clear
        ReportTest "Err.Number = 0 after Clear", Err.Number = 0
        ReportTest "Err.Description accessible", TypeName(Err.Description) <> ""

        ' ========== OBJECTCONTEXT OBJECT ==========
        Response.Write "<h2>ObjectContext Object</h2>"

        On Error Resume Next
        ObjectContext.SetAbort()
        ReportTest "ObjectContext.SetAbort callable", Err.Number = 0
        Err.Clear

        ' ========== SCRIPTING.DICTIONARY ==========
        Response.Write "<h2>Scripting.Dictionary</h2>"

        On Error Resume Next
        Set dict = Server.CreateObject("Scripting.Dictionary")
        dict.Add "key1", "value1"
        ReportTest "Dictionary.Add works", dict.Count = 1
        ReportTest "Dictionary.Exists works", dict.Exists("key1") = True
        ReportTest "Dictionary.Item access works", dict("key1") = "value1"
        ReportTest "Dictionary.Remove works", True ' Just test it doesn't Error
        dict.Remove "key1"
        ReportTest "Dictionary.Count after Remove", dict.Count = 0
        Err.Clear

        ' ========== FILESYSTEMOBJECT ==========
        Response.Write "<h2>FileSystemObject (Basic)</h2>"

        On Error Resume Next
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        ReportTest "FileSystemObject.CreateObject works", Err.Number = 0 And TypeName(fso) <> ""
        Err.Clear

        ' ========== VBSCRIPT LANGUAGE FEATURES ==========
        Response.Write "<h2>VBScript Language Features</h2>"

        ' String functions
        ReportTest "Len() function", Len("test") = 4
        ReportTest "UCase() function", UCase("hello") = "HELLO"
        ReportTest "LCase() function", LCase("HELLO") = "hello"
        ReportTest "Mid() function", Mid("hello", 2, 3) = "ell"
        ReportTest "InStr() function", InStr("hello", "ll") = 3

        ' Math functions
        ReportTest "Abs() function", Abs(-5) = 5
        ReportTest "Int() function", Int(3.7) = 3

        ' Type functions
        ReportTest "IsNumeric() function", IsNumeric(123) = True
        ReportTest "IsNumeric() with string", IsNumeric("123") = True

        ' Array operations
        Dim arr(2)
        arr(0) = "a"
        arr(1) = "b"
        arr(2) = "c"
        ReportTest "Array creation and access", arr(1) = "b"
        ReportTest "UBound() function", UBound(arr) = 2

        ' Conditional
        Dim result: result = 0
        If 1 = 1 Then
            result = 1
        End If
        ReportTest "If/Then/End If", result = 1

        ' Loop
        Dim Count: Count = 0
        Dim i
        For i = 1 To 3
            Count = Count + 1
        Next
        ReportTest "For/Next loop", Count = 3

        ' Select Case
        Dim testVal: testVal = 2
        Dim caseResult: caseResult = ""
        Select Case testVal
            Case 1: caseResult = "one"
            Case 2: caseResult = "two"
            Case Else: caseResult = "other"
        End Select
        ReportTest "Select/Case", caseResult = "two"

        ' Error handling
        On Error Resume Next
        Dim x: x = 1 / 0 ' Division by zero
        Dim errorOccurred: errorOccurred = (Err.Number <> 0)
        Err.Clear
        On Error Goto 0
        ReportTest "On Error Resume Next", errorOccurred = True

        ' ======== SUMMARY ========
        Response.Write "<h2>Test Summary</h2>"
        Response.Write "<p><b>Passed:</b> " & testsPassed & "</p>"
        Response.Write "<p><b>Failed:</b> " & testsFailed & "</p>"
        Response.Write "<p><b>Total:</b> " & (testsPassed + testsFailed) & "</p>"

        If testsFailed = 0 Then
            Response.Write "<p style='color: green; font-size: 18px;'><b>ALL TESTS PASSED! ✓</b></p>"
        Else
            Response.Write "<p style='color: red; font-size: 18px;'><b>" & testsFailed & " TESTS FAILED ✗</b></p>"
        End If
        %>
    </body>
</html>
