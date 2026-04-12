<%
' Test function returning objects
On Error Resume Next

Class Simple
    Public Value
End Class

Function ReturnObject()
    Dim obj
    Set obj = New Simple
    obj.Value = "test"
    Set ReturnObject = obj
End Function

Dim result
Set result = ReturnObject()
Response.Write "result.Value = " & result.Value & vbCrLf
%>
