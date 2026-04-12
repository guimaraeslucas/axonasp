<%
Class Test
    Function GetString(Input)
        GetString = Input
    End Function
End Class

Dim t
Set t = New Test

Dim v
v = t.GetString("hello")

Response.Write v
%>
