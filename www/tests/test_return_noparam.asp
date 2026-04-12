<%
Class Inner
    Public Value
End Class

Class Test
    Function GetInner(param)
        Dim obj
        Set obj = New Inner
        obj.Value = "fixed"
        Set GetInner = obj
    End Function
End Class

Dim t
Set t = New Test

Dim result
Set result = t.GetInner("test")

Response.Write "Type: " & TypeName(result) & vbCrLf
%>
