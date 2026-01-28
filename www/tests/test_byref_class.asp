<%
Option Explicit
Response.ContentType = "text/plain"

Class TestByRef
    Public Sub ModifyVar(x)
        x = x + 10
    End Sub
    
    Public Function TestIt()
        Dim myVar
        myVar = 5
        ModifyVar myVar
        TestIt = myVar
    End Function
End Class

Dim obj
Set obj = New TestByRef
Dim result
result = obj.TestIt()

Response.Write "Result: " & result & vbCrLf
Response.Write "Expected: 15" & vbCrLf
If result = 15 Then
    Response.Write "PASS - ByRef working in class methods"
Else
    Response.Write "FAIL - ByRef NOT working (got " & result & ")"
End If
%>
