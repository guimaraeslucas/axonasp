<%
Option Explicit
Response.ContentType = "text/plain"

Class TestDebug
    Public Sub ModifyVar(x)
        Response.Write "In ModifyVar, x before: " & x & vbCrLf
        x = x + 10
        Response.Write "In ModifyVar, x after: " & x & vbCrLf
    End Sub
    
    Public Function TestIt()
        Dim myVar
        myVar = 5
        Response.Write "Before ModifyVar, myVar: " & myVar & vbCrLf
        ModifyVar myVar
        Response.Write "After ModifyVar, myVar: " & myVar & vbCrLf
        TestIt = myVar
    End Function
End Class

Dim obj
Set obj = New TestDebug
Dim result
result = obj.TestIt()

Response.Write vbCrLf & "Final Result: " & result & vbCrLf
If result = 15 Then
    Response.Write "PASS"
Else
    Response.Write "FAIL (expected 15, got " & result & ")"
End If
%>
