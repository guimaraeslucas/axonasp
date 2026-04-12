<%
Class Inner
    Public Value
End Class

Class Test
    Function GetInner(param)        ' With parameter
        Dim obj
        Set obj = New Inner
        obj.Value = param
        Set GetInner = obj
    End Function
End Class

Dim t
Set t = New Test

Dim result
Set result = t.GetInner("test")

Response.Write "Type: " & TypeName(result) & vbCrLf
Response.Write "Value: " & result.Value & vbCrLf
%>
