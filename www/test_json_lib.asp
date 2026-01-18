<%@ Language="VBScript" %>
<html>
<head><title>JSON Library Test</title></head>
<body>

<%
    ' Test JSON library
    Set json = Server.CreateObject("G3JSON")
    Response.Write("<p>JSON Object created: " & TypeName(json) & "</p>")
    
    ' Try to use it
    obj = json.NewObject()
    Response.Write("<p>NewObject result: " & TypeName(obj) & "</p>")
%>

</body>
</html>
