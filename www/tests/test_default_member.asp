<%
' Test default member call
Class TestClass
    Public Default Sub Exec(path)
        Response.Write "Exec called with: " & path & "<br>"
    End Sub
End Class

Dim obj
Set obj = New TestClass

' Direct call using default member
Response.Write "Calling obj with argument:<br>"
obj("test.inc")

Response.Write "<br>Test complete"
%>
