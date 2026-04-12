<%
Response.Write "Testing Private Default Function<br><hr><br>"

On Error Resume Next

Class TestPrivateDefault
    Private Default Function GetValue
        GetValue = "Private Default Works!"
    End Function
End Class

Class TestPublicDefault
    Public Default Function GetValue
        GetValue = "Public Default Works!"
    End Function
End Class

' Test Public Default
Dim obj1
Set obj1 = New TestPublicDefault
If Err.Number <> 0 Then
    Response.Write "Error creating TestPublicDefault: " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "<b>TestPublicDefault created successfully</b><br>"
    Dim result1
    result1 = obj1()
    If Err.Number <> 0 Then
        Response.Write "Error calling default: " & Err.Description & "<br>"
        Err.Clear
    Else
        Response.Write "Public Default result: " & result1 & "<br>"
    End If
End If

Response.Write "<br>"

' Test Private Default
Dim obj2
Set obj2 = New TestPrivateDefault
If Err.Number <> 0 Then
    Response.Write "Error creating TestPrivateDefault: " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "<b>TestPrivateDefault created successfully</b><br>"
    Dim result2
    result2 = obj2()
    If Err.Number <> 0 Then
        Response.Write "Error calling default: " & Err.Description & "<br>"
        Err.Clear
    Else
        Response.Write "Private Default result: " & result2 & "<br>"
    End If
End If

Response.Write "<br><hr><br><b>Compilation Test: PASSED</b><br>"
Response.Write "Both Public Default and Private Default functions compiled successfully."
%>
