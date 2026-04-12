<%
' Tests eval("new ClassName") inside a class method that also calls executeGlobal

Class Plugin
    Public path
    Private p_conn

    Private Sub Class_Initialize()
        path = ""
        Set p_conn = Nothing
    End Sub

    Public Function getConn()
        Set getConn = p_conn
    End Function
End Class

Class App
    Private plugins

    Private Sub Class_Initialize()
        Set plugins = Nothing
    End Sub

    Public Function dict()
        Set dict = Server.CreateObject("Scripting.Dictionary")
    End Function

    ' Loads a plugin by executing code that defines a class, then returns an instance via eval
    Public Default Sub exec(code)
        On Error Resume Next
        executeGlobal code
        If Err.number <> 0 Then
            Response.Write "exec error: " & Err.description & vbCrLf
        End If
        On Error Goto 0
    End Sub

    Public Function plugin(name)
        name = LCase(name)
        Response.Write "plugins is nothing: " & (plugins Is Nothing) & vbCrLf
        If plugins Is Nothing Then Set plugins = dict
        Response.Write "plugins type after dict: " & TypeName(plugins) & vbCrLf

        If Not plugins.Exists(name) Then
            Dim pluginCode : pluginCode = "class cls_app_" & name & vbCrLf & _
                                          "  public path" & vbCrLf & _
                                          "  Private Sub Class_Initialize()" & vbCrLf & _
                                          "    path = """"" & vbCrLf & _
                                          "  End Sub" & vbCrLf & _
                                          "End Class"
            exec(pluginCode)
            plugins.Add name, ""
        End If

        Set plugin = Eval("new cls_app_" & name)
        Response.Write "plugin typename: " & TypeName(plugin) & vbCrLf
    End Function
End Class

Dim app : Set app = New App
Dim db
Set db = app.plugin("database")
Response.Write "db type: " & TypeName(db) & vbCrLf
db.path = "test.mdb"
Response.Write "db.path: " & db.path & vbCrLf
Response.Write "TEST PASSED" & vbCrLf
%>
