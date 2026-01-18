<%@ Language=VBScript %>
<pre>
<%
    Dim http, todo
    Set http = Server.CreateObject("G3HTTP")
    Set todo = http.Fetch("https://jsonplaceholder.typicode.com/todos/1")
    
    Response.Write "TypeName: " & TypeName(todo) & vbCrLf
    Response.Write "IsObject: " & IsObject(todo) & vbCrLf
    
    ' Tentar acessar Item diretamente
    Response.Write "Item(""id""): " & todo.Item("id") & vbCrLf
    Response.Write "GetProperty(""id""): " & todo.GetProperty("id") & vbCrLf
%>
</pre>
