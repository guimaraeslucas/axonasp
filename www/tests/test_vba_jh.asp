<%@ Language=VBScript %>
<%
' --- Interface Definition ---
Class IFoo
    Public Sub DoSomething()
    End Sub
End Class

' --- Concrete Implementation 1 ---
Class Foo
    Implements IFoo
    Public Sub IFoo_DoSomething()
        Response.Write "Foo.DoSomething called" & vbCrLf
    End Sub
End Class

' --- Concrete Implementation 2 (Failing Context) ---
Class Bar
    Implements IFoo
    Private m_Foo As IFoo

    Public Sub Init()
        Set m_Foo = New Foo
    End Sub

    Public Sub IFoo_DoSomething()
        Response.Write "Bar.DoSomething calling m_Foo..." & vbCrLf
        m_Foo.DoSomething    ' <-- ERROR HERE: Fails to map to m_Foo.IFoo_DoSomething
        Response.Write "Bar.DoSomething done" & vbCrLf
    End Sub
End Class

' --- Execution ---
Dim barObj
Set barObj = New Bar
barObj.Init

Dim bar As IFoo
Set bar = barObj
bar.DoSomething
%>