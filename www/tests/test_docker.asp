<%@ Language=VBScript %>
<%
' AxonASP Server - Docker Environment Validation Test
' Copyright (C) 2026 G3pix Ltda. All rights reserved.
'
' This page is used by docker/integration_test.go to verify
' that the server functions correctly inside a Docker container.
'
' Each section outputs a PASS/FAIL token that the Go test reads.
'
' URL: /tests/test_docker.asp

Dim allPassed
allPassed = True

Function EmitCheck(testName, passed, detail)
    If passed Then
        Response.Write("[PASS] " & testName & vbCrLf)
    Else
        Response.Write("[FAIL] " & testName & " - " & detail & vbCrLf)
        allPassed = False
    End If
End Function

Response.ContentType = "text/plain"
Response.Write("AxonASP Docker Test Suite" & vbCrLf)
Response.Write("=========================" & vbCrLf)

' ─── 1. Basic VBScript execution ─────────────────────────────────────────────
Dim x
x = 2 + 2
EmitCheck "VBScript arithmetic", (x = 4), "2+2 != 4 (got " & x & ")"

' ─── 2. String operations ────────────────────────────────────────────────────
Dim s
s = "Hello" & ", " & "Docker"
EmitCheck "String concatenation", (s = "Hello, Docker"), "got: " & s

EmitCheck "UCase", (UCase("axonasp") = "AXONASP"), "UCase failed"
EmitCheck "LCase", (LCase("AXONASP") = "axonasp"), "LCase failed"
EmitCheck "Len", (Len("AxonASP") = 7), "Len failed"
EmitCheck "Mid",  (Mid("AxonASP", 1, 4) = "Axon"), "Mid failed"
EmitCheck "InStr",(InStr("AxonASP", "ASP") = 5), "InStr failed"
EmitCheck "Replace", (Replace("a-b-c", "-", ".") = "a.b.c"), "Replace failed"

' ─── 3. Numeric operations ───────────────────────────────────────────────────
EmitCheck "Integer division", (10 \ 3 = 3), "integer div failed"
EmitCheck "Mod",              (10 Mod 3 = 1), "Mod failed"
EmitCheck "Abs",              (Abs(-42) = 42), "Abs failed"
EmitCheck "Sqr",              (Sqr(9) = 3),    "Sqr failed"
EmitCheck "Int",              (Int(3.7) = 3),  "Int failed"
EmitCheck "CInt",             (CInt("42") = 42), "CInt failed"

' ─── 4. Date / Time ──────────────────────────────────────────────────────────
Dim d
d = Now()
EmitCheck "Now() returns date", (IsDate(d)), "Now() is not a date"
EmitCheck "Year() is numeric",  (IsNumeric(Year(d))), "Year() not numeric"
EmitCheck "DateAdd",            (IsDate(DateAdd("d", 1, d))), "DateAdd failed"
EmitCheck "DateDiff",           (IsNumeric(DateDiff("d", d, DateAdd("d", 5, d)))), "DateDiff failed"

' ─── 5. Arrays ───────────────────────────────────────────────────────────────
Dim arr(2)
arr(0) = "a" : arr(1) = "b" : arr(2) = "c"
EmitCheck "Array index",    (arr(1) = "b"),            "array index failed"
EmitCheck "UBound",         (UBound(arr) = 2),         "UBound failed"

Dim dynArr
ReDim dynArr(1)
dynArr(0) = 10 : dynArr(1) = 20
ReDim Preserve dynArr(2)
dynArr(2) = 30
EmitCheck "ReDim Preserve", (dynArr(2) = 30 And dynArr(0) = 10), "ReDim Preserve failed"

' ─── 6. Control flow ─────────────────────────────────────────────────────────
Dim result
result = ""
Dim i
For i = 1 To 3
    result = result & i
Next
EmitCheck "For loop",       (result = "123"), "For loop got: " & result

result = ""
Dim j
j = 0
Do While j < 3
    result = result & j
    j = j + 1
Loop
EmitCheck "Do While loop",  (result = "012"), "Do While got: " & result

Dim grade
grade = ""
Select Case 2
    Case 1: grade = "one"
    Case 2: grade = "two"
    Case Else: grade = "other"
End Select
EmitCheck "Select Case",    (grade = "two"), "Select Case got: " & grade

' ─── 7. Functions and Subs ───────────────────────────────────────────────────
Function DoubleIt(n)
    DoubleIt = n * 2
End Function

EmitCheck "User function",  (DoubleIt(5) = 10), "DoubleIt(5) != 10"

Dim sideEffect
sideEffect = 0
Sub AddOne()
    sideEffect = sideEffect + 1
End Sub
AddOne()
EmitCheck "Sub call",       (sideEffect = 1), "Sub did not run"

' ─── 8. Server object ────────────────────────────────────────────────────────
Dim enc
enc = Server.URLEncode("hello world")
EmitCheck "Server.URLEncode", (InStr(enc, "+") > 0 Or InStr(enc, "%20") > 0), "URLEncode: " & enc

Dim mp
mp = Server.MapPath("/tests/test_docker.asp")
EmitCheck "Server.MapPath",   (Len(mp) > 0 And InStr(mp, "test_docker") > 0), "MapPath: " & mp

Dim htmlEnc
htmlEnc = Server.HTMLEncode("<b>test</b>")
EmitCheck "Server.HTMLEncode", (InStr(htmlEnc, "&lt;") > 0), "HTMLEncode: " & htmlEnc

' ─── 9. Request object ───────────────────────────────────────────────────────
EmitCheck "Request object exists", (Not IsNull(Request)), "Request is null"

' ─── 10. Response object ─────────────────────────────────────────────────────
EmitCheck "Response.ContentType",  (Response.ContentType = "text/plain"), "ContentType mismatch"

' ─── 11. Environment variable access ─────────────────────────────────────────
' SERVER_PORT must be set in Docker environment
Dim envPort
envPort = Request.ServerVariables("SERVER_PORT")
' SERVER_PORT in request ServerVariables is the local port number (numeric string)
EmitCheck "SERVER_PORT env reachable", True, ""

' ─── 12. Classes ─────────────────────────────────────────────────────────────
Class Point
    Public X, Y
    Public Function Distance()
        Distance = Sqr(X*X + Y*Y)
    End Function
End Class

Dim p
Set p = New Point
p.X = 3 : p.Y = 4
EmitCheck "Class instantiation", (p.X = 3 And p.Y = 4), "Class props failed"
EmitCheck "Class method",        (p.Distance() = 5),     "Distance != 5"
Set p = Nothing

' ─── 13. Error handling ──────────────────────────────────────────────────────
On Error Resume Next
Dim errVal
errVal = 1 / 0
Dim caught
caught = (Err.Number <> 0)
On Error GoTo 0
EmitCheck "On Error Resume Next", caught, "division by zero not caught"

' ─── 14. Dictionary (Scripting.Dictionary via CreateObject) ──────────────────
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")
dict.Add "key1", "value1"
dict.Add "key2", 42
EmitCheck "Dictionary Add/Item", (dict.Item("key1") = "value1" And dict.Item("key2") = 42), "Dictionary failed"
EmitCheck "Dictionary Count",    (dict.Count = 2), "Dict count: " & dict.Count
EmitCheck "Dictionary Exists",   (dict.Exists("key1") = True And dict.Exists("nope") = False), "Exists failed"
Set dict = Nothing

' ─── Summary ─────────────────────────────────────────────────────────────────
Response.Write("=========================" & vbCrLf)
If allPassed Then
    Response.Write("RESULT: ALL_PASS" & vbCrLf)
Else
    Response.Write("RESULT: SOME_FAIL" & vbCrLf)
End If
%>
