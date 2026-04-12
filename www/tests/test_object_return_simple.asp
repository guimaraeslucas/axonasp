<%
Class Inner
    Public Value
End Class

Class Test
    Function GetInner()
        Dim obj
        Set obj = New Inner
        obj.Value = "hello"
        Set GetInner = obj
    End Function
End Class

Dim t
Set t = New Test

Dim result
Set result = t.GetInner()

Response.Write "Type check: " & TypeName(result) & vbCrLf
Response.Write "Value: " & result.Value & vbCrLf
%>
