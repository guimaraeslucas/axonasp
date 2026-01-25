<%@ Language=VBScript %>
<pre>
<%
    Dim http, todo
    Set http = Server.CreateObject("G3HTTP")
    Set todo = http.Fetch("https://jsonplaceholder.typicode.com/todos/1")
    
    Response.Write "TypeName: " & TypeName(todo) & vbCrLf
    Response.Write "IsObject: " & IsObject(todo) & vbCrLf
    Response.Write vbCrLf
    
    On Error Resume Next
    Dim id, title, completed, userId
    
    id = todo("id")
    Response.Write "Erro Number after todo(""id""): " & Err.Number & vbCrLf
    Response.Write "Erro Description: " & Err.Description & vbCrLf
    Response.Write "Valor de ID: " & id & vbCrLf
    Err.Clear
    
    On Error GoTo 0
%>
</pre>