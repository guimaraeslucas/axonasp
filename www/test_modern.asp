<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Modern Features Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .result { border-left: 4px solid #667eea; padding: 15px; margin: 15px 0; background: #f9f9f9; border-radius: 4px; }
        .success { color: #28a745; }
        .error { color: #dc3545; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Modern Features Test</h1>
        <div class="intro">
            <p>Tests modern ASP features like environment variables and external API consumption with Fetch.</p>
        </div>
        <div class="result">
            <%
                ' 1. Pegar configuração do ambiente
                Dim apiKey
                apiKey = Env("API_KEY") 
                Response.Write("Chave da API do Ambiente: " & apiKey & "<br><br>")
                
                ' 2. Consumir uma API real (JSONPlaceholder)
                Dim http, todo
                Set http = Server.CreateObject("G3HTTP")
                Set todo = http.Fetch("https://jsonplaceholder.typicode.com/todos/1")
                
                Response.Write("<h3>Teste de Modernização</h3>")
                
                If IsObject(todo) Then
                    Response.Write("ID: " & todo("id") & "<br>")
                    Response.Write("Título: " & todo("title") & "<br>")
                    
                    If todo("completed") Then
                        Response.Write("Status: <span style='color:green'>Concluído</span><br>")
                    Else
                        Response.Write("Status: <span style='color:red'>Pendente</span><br>")
                    End If
                Else
                    Response.Write("<span class='error'>Erro ao buscar dados. Verificar console.</span><br>")
                End If
            %>
        </div>
    </div>
</body>
</html>