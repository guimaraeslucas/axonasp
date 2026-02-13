<%
Option Explicit
Dim globalVar
globalVar = "I am Global"

Dim globalObj
Set globalObj = Server.CreateObject("Scripting.Dictionary")
globalObj.Add "Key", "Global Object Value"

Class TestClass
    Public Function ReadGlobal()
        Response.Write "Class Method Reading Global: " & globalVar & "<br>"
        If IsObject(globalObj) Then
            Response.Write "Class Method Reading Global Object: " & globalObj("Key") & "<br>"
        Else
            Response.Write "Class Method: globalObj is NOT an object<br>"
        End If
    End Function

    Public Function CheckDB()
         On Error Resume Next
         Dim conn
         Set conn = Server.CreateObject("ADODB.Connection")
         Response.Write "Created ADODB in Class: " & TypeName(conn) & "<br>"
         If Err.Number <> 0 Then Response.Write "Error creating ADODB: " & Err.Description & "<br>"
    End Function
End Class

Dim c
Set c = New TestClass
c.ReadGlobal()
c.CheckDB()

Response.Write "Test Complete<br>"
%>