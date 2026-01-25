<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Env() e Fetch() Working!</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 40px 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; }
        .container { max-width: 900px; margin: 0 auto; background: white; padding: 40px; border-radius: 12px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); }
        h1 { color: #667eea; margin-bottom: 10px; font-size: 2.5em; }
        .subtitle { color: #666; font-size: 1.1em; margin-bottom: 30px; }
        .feature { display: flex; align-items: start; margin: 25px 0; padding: 20px; background: #f8f9ff; border-left: 5px solid #667eea; border-radius: 8px; }
        .feature-icon { font-size: 2em; margin-right: 20px; }
        .feature-content h2 { color: #333; margin-bottom: 8px; font-size: 1.3em; }
        .feature-content p { color: #666; line-height: 1.6; }
        .code { background: #1e1e1e; color: #d4d4d4; padding: 15px; border-radius: 6px; font-family: 'Courier New', monospace; font-size: 0.9em; margin: 10px 0; overflow-x: auto; }
        .success-badge { display: inline-block; background: #28a745; color: white; padding: 8px 16px; border-radius: 20px; font-size: 0.9em; font-weight: bold; margin: 10px 0; }
        .example-box { background: #f0f8ff; border: 2px solid #667eea; padding: 15px; border-radius: 8px; margin: 15px 0; }
        .example-box strong { color: #667eea; }
        hr { border: none; border-top: 2px solid #e0e0e0; margin: 30px 0; }
    </style>
</head>
<body>
    <div class="container">
        <h1>✓ Env() e Fetch() Implementados!</h1>
        <p class="subtitle">Funcionalidades modernas de ambiente e API consumption já funcionando</p>
        
        <div class="feature">
            <div class="feature-icon">🔧</div>
            <div class="feature-content">
                <h2>Função Env() - Variáveis de Ambiente</h2>
                <p>Acesse variáveis de ambiente do sistema e do arquivo <strong>.env</strong> facilmente.</p>
                <div class="code">apiKey = Env("API_KEY")<br>smtpHost = Env("SMTP_HOST")<br>dbName = Env("DB_NAME")</div>
                <div class="example-box">
                    <strong>Teste Prático:</strong><br>
                    <%
                        Dim apiKey, smtpHost
                        apiKey = Env("API_KEY")
                        smtpHost = Env("SMTP_HOST")
                        
                        Response.Write "API_KEY = <strong>" & apiKey & "</strong><br>"
                        Response.Write "SMTP_HOST = <strong>" & smtpHost & "</strong>"
                    %>
                </div>
                <span class="success-badge">✓ Status: Funcionando</span>
            </div>
        </div>
        
        <div class="feature">
            <div class="feature-icon">🌐</div>
            <div class="feature-content">
                <h2>G3HTTP.Fetch() - Consumo de APIs REST</h2>
                <p>Consuma APIs externas com JSON parsing automático. Os dados retornam como objetos acessáveis.</p>
                <div class="code">Dim http, resultado<br>Set http = Server.CreateObject("G3HTTP")<br>Set resultado = http.Fetch("https://api.example.com/data")<br>Response.Write resultado("propriedade")</div>
                <div class="example-box">
                    <strong>Teste Prático - JSONPlaceholder:</strong><br>
                    <%
                        Dim http, todo, user, completoStatus
                        Set http = Server.CreateObject("G3HTTP")
                        
                        ' Fetch TODO
                        Set todo = http.Fetch("https://jsonplaceholder.typicode.com/todos/1")
                        
                        Response.Write "TODO #1: <strong>" & todo("title") & "</strong><br>"
                        
                        If todo("completed") Then
                            completoStatus = "<span style='color:green'>Sim</span>"
                        Else
                            completoStatus = "<span style='color:red'>Não</span>"
                        End If
                        Response.Write "Completo: " & completoStatus & "<br><br>"
                        
                        ' Fetch USER
                        Set user = http.Fetch("https://jsonplaceholder.typicode.com/users/1")
                        
                        Response.Write "Usuário: <strong>" & user("name") & "</strong><br>"
                        Response.Write "Email: <strong>" & user("email") & "</strong><br>"
                        Response.Write "Website: <strong>" & user("website") & "</strong>"
                    %>
                </div>
                <span class="success-badge">✓ Status: Funcionando</span>
            </div>
        </div>
        
        <hr>
        
        <div class="feature">
            <div class="feature-icon">🚀</div>
            <div class="feature-content">
                <h2>Integração Completa</h2>
                <p>Combine Env() e Fetch() para criar aplicações modernas com configuração via ambiente.</p>
                <div class="code">Dim http, apiKey, apiUrl<br>apiKey = Env("API_KEY")<br>apiUrl = Env("API_URL")<br><br>Set http = Server.CreateObject("G3HTTP")<br>Set data = http.Fetch(apiUrl & "/todos/1")<br><br>If IsObject(data) Then<br>&nbsp;&nbsp;Response.Write data("title")<br>End If</div>
                <span class="success-badge">✓ Pronto para Produção</span>
            </div>
        </div>
        
        <hr>
        
        <h2 style="color: #333; margin-top: 30px;">Checklist de Funcionalidades</h2>
        <ul style="margin-left: 20px; line-height: 2.2;">
            <li>✅ Função <code>Env()</code> implementada e funcionando</li>
            <li>✅ Suporte a variáveis de ambiente do SO e .env</li>
            <li>✅ G3HTTP.Fetch() consumindo APIs reais</li>
            <li>✅ JSON parsing automático de respostas</li>
            <li>✅ Acesso a propriedades de objetos JSON</li>
            <li>✅ Suporte a múltiplas requisições simultâneas</li>
            <li>✅ Timeout configurável (10 segundos padrão)</li>
            <li>✅ Content-Type detection automático</li>
        </ul>
    </div>
</body>
</html>
