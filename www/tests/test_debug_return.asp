<%
' Debug: Very simple class method return
Class Test
    Function Ret()
        Ret = "xyz"
    End Function
End Class

Dim t
Set t = New Test

Dim v
v = t.Ret()

Response.Write v
%>
