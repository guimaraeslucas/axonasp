<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Variáveis de Ambiente & API Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1200px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 3px solid #667eea; padding-bottom: 10px; }
        h2 { color: #666; margin-top: 25px; margin-bottom: 15px; border-left: 4px solid #667eea; padding-left: 15px; }
        h3 { color: #888; margin-top: 15px; margin-bottom: 10px; }
        .section { border: 1px solid #e0e0e0; padding: 20px; margin: 20px 0; border-radius: 6px; background: #fafafa; }
        .code-block { background: #f4f4f4; padding: 15px; border-left: 3px solid #667eea; font-family: monospace; margin: 10px 0; overflow-x: auto; }
        .success { color: #28a745; font-weight: bold; }
        .error { color: #dc3545; font-weight: bold; }
        .info { color: #0066cc; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        table th { background: #f0f0f0; padding: 10px; text-align: left; border-bottom: 2px solid #667eea; }
        table td { padding: 10px; border-bottom: 1px solid #e0e0e0; }
        table tr:hover { background: #f9f9f9; }
        .data-display { background: white; border: 1px solid #ddd; padding: 15px; border-radius: 4px; margin: 15px 0; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Env() e API Fetch Test</h1>
        
        <div class="section">
            <h2>1. Função Env() - Variáveis de Ambiente</h2>
            <p>A função <code>Env()</code> permite acessar variáveis de ambiente do sistema e do arquivo <code>.env</code>.</p>
            
            <h3>Sintaxe:</h3>
            <div class="code-block">valor = Env("NOME_DA_VARIAVEL")</div>
            
            <h3>Teste Prático:</h3>
            <table>
                <tr>
                    <th>Variável</th>
                    <th>Valor</th>
                    <th>Status</th>
                </tr>
                <%
                    Dim apiKey, smtpHost, dbName
                    
                    apiKey = Env("API_KEY")
                    smtpHost = Env("SMTP_HOST")
                    dbName = Env("DB_NAME")
                    
                    ' API_KEY
                    Response.Write "<tr>"
                    Response.Write "<td>API_KEY</td>"
                    Response.Write "<td>" & apiKey & "</td>"
                    If apiKey <> "" Then
                        Response.Write "<td><span class='success'>✓ Encontrada</span></td>"
                    Else
                        Response.Write "<td><span class='error'>✗ Não definida</span></td>"
                    End If
                    Response.Write "</tr>"
                    
                    ' SMTP_HOST
                    Response.Write "<tr>"
                    Response.Write "<td>SMTP_HOST</td>"
                    Response.Write "<td>" & smtpHost & "</td>"
                    If smtpHost <> "" Then
                        Response.Write "<td><span class='success'>✓ Encontrada</span></td>"
                    Else
                        Response.Write "<td><span class='error'>✗ Não definida</span></td>"
                    End If
                    Response.Write "</tr>"
                    
                    ' DB_NAME
                    Response.Write "<tr>"
                    Response.Write "<td>DB_NAME</td>"
                    Response.Write "<td>" & dbName & "</td>"
                    If dbName <> "" Then
                        Response.Write "<td><span class='success'>✓ Encontrada</span></td>"
                    Else
                        Response.Write "<td><span class='error'>✗ Não definida</span></td>"
                    End If
                    Response.Write "</tr>"
                %>
            </table>
        </div>
        
        <div class="section">
            <h2>2. G3HTTP.Fetch() - Consumo de API</h2>
            <p>O método <code>Fetch()</code> permite consumir APIs REST e retorna automaticamente JSON parseado como objeto.</p>
            
            <h3>Sintaxe:</h3>
            <div class="code-block">
Dim http<br>
Set http = Server.CreateObject("G3HTTP")<br>
Set resultado = http.Fetch("https://api.exemplo.com/endpoint")<br>
Response.Write resultado("propriedade")
            </div>
            
            <h3>Teste Prático - JSONPlaceholder API:</h3>
            <%
                Dim http, todo, user, posts
                Set http = Server.CreateObject("G3HTTP")
                
                ' Teste 1: Fetch um TODO
                Response.Write "<h4>Teste 1: Buscar um TODO (ID: 1)</h4>"
                Set todo = http.Fetch("https://jsonplaceholder.typicode.com/todos/1")
                
                If IsObject(todo) Then
                    Response.Write "<div class='data-display'>"
                    Response.Write "ID: <strong>" & todo("id") & "</strong><br>"
                    Response.Write "Usuário: <strong>" & todo("userId") & "</strong><br>"
                    Response.Write "Título: <strong>" & todo("title") & "</strong><br>"
                    If todo("completed") Then
                        Response.Write "Status: <span class='success'>✓ Concluído</span>"
                    Else
                        Response.Write "Status: <span class='error'>✗ Pendente</span>"
                    End If
                    Response.Write "</div>"
                    Response.Write "<p class='success'>✓ Requisição bem-sucedida e JSON parseado corretamente</p>"
                Else
                    Response.Write "<p class='error'>✗ Falha ao buscar dados</p>"
                End If
                
                ' Teste 2: Fetch um USER
                Response.Write "<h4 style='margin-top: 25px;'>Teste 2: Buscar um USER (ID: 1)</h4>"
                Set user = http.Fetch("https://jsonplaceholder.typicode.com/users/1")
                
                If IsObject(user) Then
                    Response.Write "<div class='data-display'>"
                    Response.Write "ID: <strong>" & user("id") & "</strong><br>"
                    Response.Write "Nome: <strong>" & user("name") & "</strong><br>"
                    Response.Write "Email: <strong>" & user("email") & "</strong><br>"
                    Response.Write "Telefone: <strong>" & user("phone") & "</strong><br>"
                    Response.Write "Website: <strong>" & user("website") & "</strong>"
                    Response.Write "</div>"
                    Response.Write "<p class='success'>✓ Requisição bem-sucedida e JSON parseado corretamente</p>"
                Else
                    Response.Write "<p class='error'>✗ Falha ao buscar dados</p>"
                End If
            %>
        </div>
        
        <div class="section">
            <h2>3. Combinando Env() e Fetch()</h2>
            <p>Você pode usar variáveis de ambiente para armazenar URLs base de APIs e chaves de autenticação.</p>
            
            <h3>Exemplo:</h3>
            <div class="code-block">
Dim http, apiKey, apiUrl, resultado<br>
apiKey = Env("API_KEY")<br>
apiUrl = Env("API_URL")<br>
<br>
Set http = Server.CreateObject("G3HTTP")<br>
Set resultado = http.Fetch(apiUrl & "/endpoint")<br>
If IsObject(resultado) Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;Response.Write resultado("data")<br>
End If
            </div>
            
            <h3>Uso Prático:</h3>
            <%
                Dim baseUrl
                baseUrl = Env("API_URL")
                
                If baseUrl = "" Then
                    baseUrl = "https://jsonplaceholder.typicode.com"
                End If
                
                Response.Write "<p><span class='info'>ℹ API Base URL:</span> <code>" & baseUrl & "</code></p>"
                Response.Write "<p><span class='success'>✓ Sistema pronto para consumir APIs usando configurações de ambiente</span></p>"
            %>
        </div>
        
        <div class="section">
            <h2>Resumo das Funcionalidades</h2>
            <ul style="line-height: 2; margin-left: 20px;">
                <li><span class="success">✓</span> Função <code>Env()</code> - Obtém variáveis de ambiente</li>
                <li><span class="success">✓</span> G3HTTP - Biblioteca para consumo de APIs REST</li>
                <li><span class="success">✓</span> JSON Parsing Automático - APIs retornam objetos parseados</li>
                <li><span class="success">✓</span> Acesso a Propriedades - objeto("propriedade") funciona perfeitamente</li>
                <li><span class="success">✓</span> Integração completa - Env() + Fetch() + JSON = API consumption moderna</li>
            </ul>
        </div>
    </div>
</body>
</html>
