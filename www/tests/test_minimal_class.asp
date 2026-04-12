<%
Class TestClass
    Public propertyName
    Public Property Get Name
        Name = propertyName
    End Property
    Public Property Let Name(val)
        propertyName = val
    End Property
End Class

Dim obj
Set obj = New TestClass
obj.Name = "Hello"
Response.Write "Result: " & obj.Name
%>
