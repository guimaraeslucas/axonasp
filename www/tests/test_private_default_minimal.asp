<%
' Minimal test for Private Default parsing

Class TestClass
    Private Default Function GetValue
        GetValue = "It works"
    End Function
End Class

Response.Write "If you see this, parsing succeeded!"
%>
