<%
@ Language = VBScript
%>
<%
debug_axonvm = "TRUE"
%>
<%
Option Explicit
%>
<%
Class TestArrayClass
    Dim myArray

    Private Sub Class_Initialize()
        ReDim myArray(2)
        myArray(0) = "First"
        myArray(1) = "Second"
        myArray(2) = "Third"
    End Sub

    Public Function GetItem(i)
        GetItem = myArray(i)
    End Function
End Class

Dim obj
Set obj = New TestArrayClass
Response.Write "Item 0: " & obj.GetItem(0) & "<br>"
Response.Write "Item 1: " & obj.GetItem(1) & "<br>"
Response.Write "Item 2: " & obj.GetItem(2) & "<br>"
%>
