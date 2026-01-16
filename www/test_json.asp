<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - JSON Operations Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2, h3 { color: #555; margin-top: 20px; margin-bottom: 15px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .code-block { background: #f4f4f4; padding: 10px; margin: 10px 0; border-radius: 4px; word-break: break-all; }
        strong { color: #333; }
        ul { margin-left: 20px; }
        li { margin: 5px 0; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - JSON Operations Test</h1>
        <div class="intro">
            <p>Tests JSON object creation, serialization, parsing and iteration.</p>
        </div>
    <%
    ' 0. Instanciar Biblioteca
    Dim json
    Set json = Server.CreateObject("G3JSON")

    ' 1. Criar um objeto novo
    Dim user
    Set user = json.NewObject()
    
    ' 2. Editar propriedades (Sintaxe Python/Go-like adaptada pro ASP)
    user("name") = "Lucas"
    user("role") = "Developer"
    user("age") = 32
    
    Response.Write("User Name: " & user("name") & "<br>")
    
    ' 3. Serializar
    Dim jsonStr
    jsonStr = json.Stringify(user)
    Response.Write("<strong>JSON Output:</strong> " & jsonStr & "<br>")
    
    ' 4. Parsear uma string complexa
    Dim data
    data = json.Parse("{""clients"": [""Google"", ""Microsoft"", ""SpaceX""], ""active"": true}")
    
    Response.Write("Active: " & data("active") & "<br>")
    
    ' 5. Iterar sobre Array JSON (Usa sua correção do For Each!)
    Response.Write("<h3>Clients List:</h3>")
    
    Dim clientList
    ' No ASP/VBScript original isso seria complexo, aqui é nativo
    clientList = data("clients") 
    
    For Each client In clientList
        Response.Write("- " & client & "<br>")
    Next
    %>
    </div>
</body>
</html>