<%
Class MyClass
    Public MyProp
    Private MyPrivate
    
    Private Sub Class_Initialize()
        MyProp = "Initialized"
        MyPrivate = "Secret"
    End Sub
    
    Public Function GetPrivate()
        GetPrivate = MyPrivate
    End Function
    
    Public Sub SayHello(name)
        Response.Write "Hello " & name & "<br>"
    End Sub
End Class

Dim obj
Set obj = New MyClass

Response.Write "Prop: " & obj.MyProp & "<br>"
Response.Write "Private: " & obj.GetPrivate() & "<br>"
obj.SayHello "World"

With obj
    .MyProp = "Changed"
    Response.Write "With Prop: " & .MyProp & "<br>"
    .SayHello "With"
End With
%>
