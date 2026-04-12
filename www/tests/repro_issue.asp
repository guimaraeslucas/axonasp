<%
Class TestClass
    Public Function CreateSomething()
        Dim obj
        Set obj = Server.CreateObject("Scripting.Dictionary")
        CreateSomething = "OK"
    End Function
End Class

Dim o
Set o = New TestClass
Response.Write "Result: " & o.CreateSomething()
%>
