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
Class TestTerminateClass
    Dim myArray

    Private Sub Class_Initialize()
        Response.Write "Initialize<br>"
        ReDim myArray(2)
        myArray(0) = "First"
        myArray(1) = "Second"
        myArray(2) = "Third"
    End Sub

    Private Sub Class_Terminate()
        Response.Write "Terminate<br>"
        Dim i
        For i = 0 To UBound(myArray)
            Response.Write "Setting item " & i & " to Nothing<br>"
            Set myArray(i) = Nothing
        Next
    End Sub

    Public Function GetItem(i)
        GetItem = myArray(i)
    End Function
End Class

Dim obj
Set obj = New TestTerminateClass
Response.Write "Item 0: " & obj.GetItem(0) & "<br>"
Set obj = Nothing
Response.Write "Done<br>"
%>
