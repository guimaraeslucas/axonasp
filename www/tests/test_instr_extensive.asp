<%
@ Language = VBScript
%>
<%
Option Explicit
Option Compare Binary

Response.ContentType = "text/plain"

Dim totalTests, passedTests, failedTests
totalTests = 0
passedTests = 0
failedTests = 0

Sub AssertEqual(testName, actualValue, expectedValue)
    totalTests = totalTests + 1
    If actualValue = expectedValue Then
        passedTests = passedTests + 1
        Response.Write "PASS | " & testName & " | expected=" & CStr(expectedValue) & " | actual=" & CStr(actualValue) & vbCrLf
    Else
        failedTests = failedTests + 1
        Response.Write "FAIL | " & testName & " | expected=" & CStr(expectedValue) & " | actual=" & CStr(actualValue) & vbCrLf
    End If
End Sub

Sub AssertError(testName, expectedErrNumber, startArg, textArg1, textArg2)
    totalTests = totalTests + 1
    Err.Clear
    On Error Resume Next
    Call InStr(startArg, textArg1, textArg2)
    If Err.Number = expectedErrNumber Then
        passedTests = passedTests + 1
        Response.Write "PASS | " & testName & " | expectedErr=" & CStr(expectedErrNumber) & " | actualErr=" & CStr(Err.Number) & vbCrLf
    Else
        failedTests = failedTests + 1
        Response.Write "FAIL | " & testName & " | expectedErr=" & CStr(expectedErrNumber) & " | actualErr=" & CStr(Err.Number) & vbCrLf
    End If
    On Error Goto 0
End Sub

Response.Write "=== InStr Extensive Compatibility Audit ===" & vbCrLf
Response.Write "Option Compare: Binary" & vbCrLf
Response.Write vbCrLf

' Basic ASCII behavior
Call AssertEqual("ASCII basic search", InStr("banana", "na"), 3)
Call AssertEqual("ASCII start in middle", InStr(4, "banana", "na"), 5)
Call AssertEqual("ASCII start after first match", InStr(2, "banana", "ba"), 0)
Call AssertEqual("ASCII token not found", InStr("banana", "zz"), 0)
Call AssertEqual("ASCII exact full string", InStr("banana", "banana"), 1)
Call AssertEqual("ASCII start beyond length", InStr(100, "banana", "na"), 0)
Call AssertEqual("ASCII match at end", InStr("banana", "a"), 2)
Call AssertEqual("ASCII match at end with start", InStr(6, "banana", "a"), 6)

' Empty pattern behavior
Call AssertEqual("Empty find with default start", InStr("abc", ""), 1)
Call AssertEqual("Empty find with explicit start", InStr(3, "abc", ""), 3)
Call AssertEqual("Empty source and empty find", InStr("", ""), 1)
Call AssertEqual("Empty source non-empty find", InStr("", "a"), 0)

' Compare argument behavior (Binary/Text)
Call AssertEqual("Compare binary case-sensitive", InStr(1, "ABC", "a", 0), 0)
Call AssertEqual("Compare text case-insensitive", InStr(1, "ABC", "a", 1), 1)
Call AssertEqual("Compare text with start", InStr(2, "AaaA", "a", 1), 2)
Call AssertEqual("Compare binary with start", InStr(2, "AaaA", "A", 0), 4)

' UTF-8 / Unicode behavior with multi-byte characters
Dim u1
u1 = "áábç漢字🙂z"
Call AssertEqual("Unicode first accented char", InStr(u1, "á"), 1)
Call AssertEqual("Unicode second accented char via start", InStr(2, u1, "á"), 2)
Call AssertEqual("Unicode cedilla", InStr(u1, "ç"), 4)
Call AssertEqual("Unicode CJK first", InStr(u1, "漢"), 5)
Call AssertEqual("Unicode CJK second", InStr(u1, "字"), 6)
Call AssertEqual("Unicode emoji", InStr(u1, "🙂"), 7)
Call AssertEqual("Unicode ASCII tail", InStr(u1, "z"), 8)
Call AssertEqual("Unicode start after codepoint", InStr(7, u1, "字"), 0)
Call AssertEqual("Unicode start at emoji", InStr(7, u1, "🙂"), 7)
Call AssertEqual("Unicode start at tail", InStr(8, u1, "z"), 8)

' Text compare with Unicode casing
Call AssertEqual("Unicode text compare fold", InStr(1, "Árvore", "á", 1), 1)
Call AssertEqual("Unicode binary compare no fold", InStr(1, "Árvore", "á", 0), 0)

' Expected runtime error behavior in Classic ASP when start < 1
' In ASP, Err.Number is exposed as HRESULT for code 5: 0x800A0005 = -2146828283
Call AssertError("Error when start = 0", -2146828283, 0, "abc", "a")
Call AssertError("Error when start < 0", -2146828283, -1, "abc", "a")

Response.Write vbCrLf
Response.Write "=== Summary ===" & vbCrLf
Response.Write "TOTAL=" & CStr(totalTests) & vbCrLf
Response.Write "PASSED=" & CStr(passedTests) & vbCrLf
Response.Write "FAILED=" & CStr(failedTests) & vbCrLf

If failedTests = 0 Then
    Response.Write "RESULT=PASS" & vbCrLf
Else
    Response.Write "RESULT=FAIL" & vbCrLf
End If
%>
