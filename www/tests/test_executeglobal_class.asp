<%
' Test: executeGlobal with class member chain calls
Class Inner
    Public Sub greet(val)
        Response.Write "Inner: " & val & vbCrLf
    End Sub
End Class

Class Outer
    Private p_inner

    Private Sub Class_Initialize()
        Set p_inner = Nothing
    End Sub

    Public Function inner()
        If p_inner Is Nothing Then
            Set p_inner = New Inner
        End If
        Set inner = p_inner
    End Function

    Public Sub showError(msg)
        inner.greet(msg)
    End Sub

    Public Default Sub exec(code)
        On Error Resume Next
        executeGlobal code
        If Err.number <> 0 Then
            showError "Error in exec: " & Err.description
        End If
        On Error Goto 0
    End Sub
End Class

Dim aspL : Set aspL = New Outer

' Simulate calling aspL("some code") which executes via executeGlobal
aspL("response.write ""From executeGlobal"" & vbCrLf")
Response.Write "After exec call" & vbCrLf

' Test 2: error in executeGlobal triggers showError (inner.greet from within exec error handler)
aspL("dim x : x = 1/0")
Response.Write "After exec error" & vbCrLf
%>
