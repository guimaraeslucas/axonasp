<%
' Test class method returning objects
On Error Resume Next

Class Inner
    Public SQL
End Class

Class Outer
    Function GetInner(cmd)
        Dim obj
        Set obj = New Inner
        obj.SQL = cmd
        Set GetInner = obj
    End Function
End Class

Dim outer
Set outer = New Outer

Dim result
Set result = outer.GetInner("select *")
Response.Write "result.SQL = " & result.SQL & vbCrLf
%>
