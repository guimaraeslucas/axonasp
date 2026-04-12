<%
Class Inner
    Public SQL
End Class

Class Test
    Function GetInner(cmd)
        Dim obj
        Set obj = New Inner
        obj.SQL = cmd
        Set GetInner = obj
    End Function
End Class

Dim t
Set t = New Test

Dim result
Set result = t.GetInner("select *")

Response.Write "Type: " & TypeName(result) & vbCrLf
Response.Write "SQL: " & result.SQL & vbCrLf
%>
