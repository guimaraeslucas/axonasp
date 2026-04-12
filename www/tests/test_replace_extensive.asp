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

Class ReplaceCarrier
    Private mValue

    Public Property Let Value(v)
        mValue = CStr(v)
    End Property

    Public Default Property Get ValueText()
        ValueText = mValue
    End Property
End Class

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

Sub AssertError(testName, expectedErrNumber, expressionArg, findArg, replaceArg, startArg, countArg, compareArg)
    totalTests = totalTests + 1
    Err.Clear

    On Error Resume Next
    If IsEmpty(startArg) Then
        Call Replace(expressionArg, findArg, replaceArg)
    ElseIf IsEmpty(countArg) Then
        Call Replace(expressionArg, findArg, replaceArg, startArg)
    ElseIf IsEmpty(compareArg) Then
        Call Replace(expressionArg, findArg, replaceArg, startArg, countArg)
    Else
        Call Replace(expressionArg, findArg, replaceArg, startArg, countArg, compareArg)
    End If

    If Err.Number = expectedErrNumber Then
        passedTests = passedTests + 1
        Response.Write "PASS | " & testName & " | expectedErr=" & CStr(expectedErrNumber) & " | actualErr=" & CStr(Err.Number) & vbCrLf
    Else
        failedTests = failedTests + 1
        Response.Write "FAIL | " & testName & " | expectedErr=" & CStr(expectedErrNumber) & " | actualErr=" & CStr(Err.Number) & vbCrLf
    End If

    On Error Goto 0
End Sub

Sub AssertIsNull(testName, valueArg)
    totalTests = totalTests + 1
    If IsNull(valueArg) Then
        passedTests = passedTests + 1
        Response.Write "PASS | " & testName & " | expected=NULL | actual=NULL" & vbCrLf
    Else
        failedTests = failedTests + 1
        Response.Write "FAIL | " & testName & " | expected=NULL | actual=" & CStr(valueArg) & vbCrLf
    End If
End Sub

Response.Write "=== Replace Extensive Compatibility Audit ===" & vbCrLf
Response.Write "Option Compare: Binary" & vbCrLf
Response.Write vbCrLf

' Basic and canonical scenarios
Call AssertEqual("Basic replace all", Replace("foo bar foo", "foo", "baz"), "baz bar baz")
Call AssertEqual("Replace not found", Replace("foo bar", "zzz", "x"), "foo bar")
Call AssertEqual("Replace to empty string", Replace("a-b-c", "-", ""), "abc")
Call AssertEqual("Replace entire input", Replace("abc", "abc", "X"), "X")
Call AssertEqual("Replace repeated char", Replace("aaaa", "a", "b"), "bbbb")

' Start argument behavior (1-based)
Call AssertEqual("Start=1 behaves as full string", Replace("abcabc", "a", "X", 1), "XbcXbc")
Call AssertEqual("Start=2 returns from start only", Replace("abcabc", "a", "X", 2), "bcXbc")
Call AssertEqual("Start in middle with no match", Replace("abcabc", "a", "X", 3), "cXbc")
Call AssertEqual("Start at string end", Replace("abcabc", "c", "Z", 6), "Z")
Call AssertEqual("Start beyond length returns empty", Replace("abcabc", "a", "X", 7), "")

' Count argument behavior
Call AssertEqual("Count=1 replaces first occurrence", Replace("foo foo foo", "foo", "bar", 1, 1), "bar foo foo")
Call AssertEqual("Count=2 replaces two occurrences", Replace("foo foo foo", "foo", "bar", 1, 2), "bar bar foo")
Call AssertEqual("Count=-1 replaces all from start", Replace("foo foo foo", "foo", "bar", 1, -1), "bar bar bar")
Call AssertEqual("Count=0 returns start-sliced string", Replace("foo foo foo", "foo", "bar", 2, 0), "oo foo foo")
Call AssertEqual("Count with start offset", Replace("foo foo foo", "foo", "bar", 2, 1), "oo bar foo")

' Compare argument behavior
Call AssertEqual("Binary compare case-sensitive", Replace("AaA", "a", "x", 1, -1, 0), "AxA")
Call AssertEqual("Text compare case-insensitive", Replace("AaA", "a", "x", 1, -1, 1), "xxx")
Call AssertEqual("Text compare with count", Replace("AaAa", "a", "x", 1, 2, 1), "xxAa")
Call AssertEqual("Binary compare with start", Replace("AaAa", "A", "x", 2, -1, 0), "axa")
Call AssertEqual("Compare=vbUseCompareOption under Binary", Replace("AaA", "a", "x", 1, -1, -1), "AxA")

' Unicode and multi-byte behavior
Dim u1
u1 = "áá漢🙂aá"
Call AssertEqual("Unicode replace all accent", Replace(u1, "á", "X"), "XX漢🙂aX")
Call AssertEqual("Unicode replace from start=2", Replace(u1, "á", "X", 2), "X漢🙂aX")
Call AssertEqual("Unicode replace CJK", Replace(u1, "漢", "字"), "áá字🙂aá")
Call AssertEqual("Unicode replace emoji", Replace(u1, "🙂", "!"), "áá漢!aá")
Call AssertEqual("Unicode text compare fold", Replace("Árvore", "á", "A", 1, -1, 1), "Arvore")
Call AssertEqual("Unicode binary compare no fold", Replace("Árvore", "á", "A", 1, -1, 0), "Árvore")

' Empty find behavior
Call AssertEqual("Empty find default start", Replace("abc", "", "X"), "abc")
Call AssertEqual("Empty find with start", Replace("abc", "", "X", 2), "bc")

' Variant coercion and object/default property behavior
Call AssertEqual("Numeric expression coerced", Replace(123123, "23", "XX"), "1XX1XX")
Call AssertEqual("Numeric find coerced", Replace("abc1abc", 1, "X"), "abcXabc")
Call AssertEqual("Numeric replace coerced", Replace("abc1abc", "1", 9), "abc9abc")

Dim carrier
Set carrier = New ReplaceCarrier
carrier.Value = "foo bar foo"
Call AssertEqual("Object explicit property string", Replace(carrier.ValueText, "foo", "baz"), "baz bar baz")

' Null behavior (classic VBScript: Replace(Null, ...) returns Null)
Call AssertIsNull("Null expression returns Null", Replace(Null, "a", "b"))

' Error behavior
' In ASP, Err.Number is exposed as HRESULT for code 5: 0x800A0005 = -2146828283
Call AssertError("Error when start=0", -2146828283, "abc", "a", "b", 0, Empty, Empty)
Call AssertError("Error when start<0", -2146828283, "abc", "a", "b", -1, Empty, Empty)

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
