<%
Class Test
    Function GetObj(cmd)
        Dim obj
        Set obj = New Test
        GetObj = obj
    End Function
End Class

Dim t
Set t = New Test
Dim result
Set result = t.GetObj("test")

Response.Write TypeName(result)
%>
