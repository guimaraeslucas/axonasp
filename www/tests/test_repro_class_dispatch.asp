<%
Response.Write "<h3>Starting Dispatch Test</h3>"

Class TestClass
    Public Default Sub DefaultSub(val)
        Response.Write "PASS: DefaultSub called with: " & val & "<br>"
    End Sub

    Public Function [isEmpty](val)
        Response.Write "PASS: isEmpty method called with: " & val & "<br>"
        isEmpty = "ReturnedFromMethod"
    End Function

    Public Sub exec(val)
        Response.Write "PASS: exec Sub called with: " & val & "<br>"
    End Sub
End Class

Dim t
Set t = New TestClass

Response.Write "<h4>1. Testing Explicit Call t.exec(""foo"")</h4>"
On Error Resume Next
t.exec("foo")
If Err.Number <> 0 Then Response.Write "FAIL: Error calling t.exec: " & Err.Description & "<br>"
On Error Goto 0

Response.Write "<h4>2. Testing Default Call t(""bar"")</h4>"
On Error Resume Next
t("bar")
If Err.Number <> 0 Then Response.Write "FAIL: Error calling t(""bar""): " & Err.Description & "<br>"
On Error Goto 0

Response.Write "<h4>3. Testing Keyword Method t.isEmpty(""baz"")</h4>"
On Error Resume Next
Dim res
res = t.isEmpty("baz")
Response.Write "Result: " & res & "<br>"
If res <> "ReturnedFromMethod" Then Response.Write "FAIL: t.isEmpty did not return expected value (Got: " & res & ")<br>"
If Err.Number <> 0 Then Response.Write "FAIL: Error calling t.isEmpty: " & Err.Description & "<br>"
On Error Goto 0

Response.Write "<h3>End Dispatch Test</h3>"
%>
