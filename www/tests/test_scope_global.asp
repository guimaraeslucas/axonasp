<%
Dim myGlobal
myGlobal = "I am Global"

Class TestGlobal
    Sub CheckGlobal()
        If myGlobal = "I am Global" Then
            Response.Write "Global Visible<br>"
        Else
            Response.Write "Global NOT Visible (" & TypeName(myGlobal) & ")<br>"
        End If
    End Sub
End Class

Dim t
Set t = New TestGlobal
t.CheckGlobal()
%>
