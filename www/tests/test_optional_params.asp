<%
Option Explicit
Response.Write("Testing OPTIONAL Parameters" & vbCrLf & vbCrLf)

' Test 1: Function with optional parameter
Function TestOptional1(required, Optional optional)
    Response.Write("required=" & required)
    If IsEmpty(optional) Then
        Response.Write(", optional=<empty>" & vbCrLf)
    Else
        Response.Write(", optional=" & optional & vbCrLf)
    End If
End Function

' Call with only required parameter
Response.Write("Call 1 (1 arg): ")
TestOptional1 "first"

' Call with both parameters
Response.Write("Call 2 (2 args): ")
TestOptional1 "first", "second"

' Test 2: Function with multiple optional parameters
Function TestOptional2(a, Optional b, Optional c, Optional d)
    Response.Write("a=" & a)
    If Not IsEmpty(b) Then Response.Write(", b=" & b)
    If Not IsEmpty(c) Then Response.Write(", c=" & c)
    If Not IsEmpty(d) Then Response.Write(", d=" & d)
    Response.Write(vbCrLf)
End Function

Response.Write(vbCrLf & "Test 2 - Multiple Optional Parameters:" & vbCrLf)
Response.Write("Call 1 (1 arg): ")
TestOptional2 "A"

Response.Write("Call 2 (2 args): ")
TestOptional2 "A", "B"

Response.Write("Call 3 (4 args): ")
TestOptional2 "A", "B", "C", "D"

' Test 3: BYREF with OPTIONAL
Sub TestOptional3(required, Optional ByRef optRef)
    Response.Write("In Sub: required=" & required)
    If IsEmpty(optRef) Then
        Response.Write(", optRef=<empty>" & vbCrLf)
        optRef = "modified"
    Else
        Response.Write(", optRef=" & optRef & vbCrLf)
        optRef = optRef & "-modified"
    End If
End Sub

Response.Write(vbCrLf & "Test 3 - BYREF with OPTIONAL:" & vbCrLf)
Dim testVar
testVar = "input"
Call TestOptional3("req", testVar)
Response.Write("After Sub: testVar=" & testVar & vbCrLf)

' Test without optional parameter
Call TestOptional3("req2")
Response.Write("After Sub (no param): test completed" & vbCrLf)

Response.Write(vbCrLf & "ALL TESTS PASSED!" & vbCrLf)
%>
