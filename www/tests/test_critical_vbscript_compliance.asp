<%
' Option Compare must be at the very top of the file
Option Compare Text

' Test 1: Option Compare Text (case-insensitive comparison)
Response.Write "<h3>Test 1: Option Compare Text</h3>"
If "ABC" = "abc" Then
    Response.Write "PASS: Strings are equal (case-insensitive)<br>"
Else
    Response.Write "FAIL: Option Compare Text not working<br>"
End If

' Test 2: ReDim Preserve
Response.Write "<h3>Test 2: ReDim Preserve</h3>"
Dim arr
arr = Array(1, 2, 3)
ReDim Preserve arr(5)
If arr(0) = 1 And UBound(arr) = 5 Then
    Response.Write "PASS: ReDim Preserve maintained original values<br>"
Else
    Response.Write "FAIL: ReDim Preserve failed<br>"
End If

' Test 3: On Error Resume Next
Response.Write "<h3>Test 3: On Error Resume Next</h3>"
On Error Resume Next
Dim x
x = 1 / 0
If Err.Number <> 0 Then
    Response.Write "PASS: Error captured (Number=" & Err.Number & ")<br>"
    Err.Clear
Else
    Response.Write "FAIL: On Error Resume Next not working<br>"
End If
On Error Goto 0

' Test 4: Eval Function
Response.Write "<h3>Test 4: Eval Function</h3>"
Dim y
y = Eval("10 + 20")
If y = 30 Then
    Response.Write "PASS: Eval returned " & y & "<br>"
Else
    Response.Write "FAIL: Eval returned " & y & " (expected 30)<br>"
End If

' Test 5: Class with Property Get/Let
Response.Write "<h3>Test 5: Class Properties</h3>"
Class TestClass
    Private pValue

    Public Property Get Value()
        Value = pValue
    End Property

    Public Property Let Value(val)
        pValue = val
    End Property
End Class

Dim obj
Set obj = New TestClass
obj.Value = 42
If obj.Value = 42 Then
    Response.Write "PASS: Class properties working<br>"
Else
    Response.Write "FAIL: Class properties failed<br>"
End If

' Test 6: Server.GetLastError
Response.Write "<h3>Test 6: Server.GetLastError</h3>"
Set e = Server.GetLastError()
If e.Number = 0 And e.Description = "" Then
    Response.Write "PASS: Server.GetLastError reports no unhandled ASP error<br>"
Else
    Response.Write "FAIL: Unexpected Server.GetLastError state: Number=" & e.Number & ", Description=" & e.Description & "<br>"
End If

Response.Write "Debug Number=" & e.Number & "<br>"
Response.Write "Debug Description=" & e.Description & "<br>"
Response.Write "Debug Source=" & e.Source & "<br>"
Response.Write "Debug Category=" & e.Category & "<br>"
Response.Write "Debug ASPCode=" & e.ASPCode & "<br>"

Response.Write "<h3>All Critical VBScript Compliance Tests Complete</h3>"
%>
