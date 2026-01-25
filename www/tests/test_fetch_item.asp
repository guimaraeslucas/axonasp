<%@ Language=VBScript %>
<pre>
<%
    Dim http, todo
    Set http = Server.CreateObject("G3HTTP")
    Set todo = http.Fetch("https://jsonplaceholder.typicode.com/todos/1")
    
    Response.Write "TypeName: " & TypeName(todo) & vbCrLf
    Response.Write "IsObject: " & IsObject(todo) & vbCrLf
    
    ' Tentar acesso direto com subscript
    On Error Resume Next
    Dim id
    id = todo("id")
    If Err.Number <> 0 Then
        Response.Write "Erro com subscript: " & Err.Description & vbCrLf
        Err.Clear
    Else
        Response.Write "todo(""id""): " & id & vbCrLf
    End If
    
    ' Tentar título
    Dim title
    title = todo("title")
    If Err.Number <> 0 Then
        Response.Write "Erro ao get title: " & Err.Description & vbCrLf
    Else
        Response.Write "todo(""title""): " & title & vbCrLf
    End If
    
    On Error GoTo 0
%>
</pre>
