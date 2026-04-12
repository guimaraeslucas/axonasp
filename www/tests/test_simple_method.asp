<%
Class Test
    Function GetString()
        GetString = "hello"
    End Function
End Class

Dim t
Set t = New Test

Dim v
v = t.GetString()

Response.Write v
%>
