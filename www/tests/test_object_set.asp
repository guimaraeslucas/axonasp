<%
' Simple test for Set assignment
On Error Resume Next

Class TestClass
    Public Name
    Function GetMe()
        GetMe = Me
    End Function
End Class

Dim obj
Set obj = New TestClass
obj.Name = "test"

Response.Write "obj.Name = " & obj.Name & vbCrLf

' Test nested function calls
Function GetObject(o)
    Set GetObject = o
End Function

Dim obj2
Set obj2 = GetObject(obj)

Response.Write "obj2.Name = " & obj2.Name & vbCrLf
%>
