<%
Class Test
    Public Value
    Function GetValue()
        GetValue = Value
    End Function
End Class

Dim t
Set t = New Test
t.Value = "myvalue"

Dim v
v = t.GetValue()

Response.Write v
%>
