<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Debug Fetch - Detalhado</title>
    <style>
        body { font-family: monospace; background: #1e1e1e; color: #d4d4d4; padding: 20px; }
        .debug-section { background: #252526; border: 1px solid #555; padding: 15px; margin: 10px 0; border-radius: 5px; }
        .label { color: #4ec9b0; font-weight: bold; }
        .value { color: #ce9178; }
        .error { color: #f48771; }
        .success { color: #6a9955; }
        h2 { color: #569cd6; border-bottom: 2px solid #569cd6; }
    </style>
</head>
<body>
    <h1>Debug Detalhado do Fetch()</h1>

    <%
        ' Teste 1: Criar objeto HTTP
        Response.Write "<div class='debug-section'>"
        Response.Write "<h2>ETAPA 1: Criar objeto G3HTTP</h2>"
        
        Dim http
        On Error Resume Next
        Set http = Server.CreateObject("G3HTTP")
        
        If Err.Number <> 0 Then
            Response.Write "<span class='error'>❌ ERRO ao criar G3HTTP: " & Err.Description & "</span><br>"
        Else
            Response.Write "<span class='success'>✓ G3HTTP criado com sucesso</span><br>"
            Response.Write "<span class='label'>Tipo:</span> <span class='value'>" & TypeName(http) & "</span><br>"
        End If
        On Error GoTo 0
        
        Response.Write "</div>"
        
        ' Teste 2: Fazer Fetch
        Response.Write "<div class='debug-section'>"
        Response.Write "<h2>ETAPA 2: Executar Fetch()</h2>"
        
        Dim url, todo
        url = "https://jsonplaceholder.typicode.com/todos/1"
        Response.Write "<span class='label'>URL:</span> <span class='value'>" & url & "</span><br>"
        
        On Error Resume Next
        Set todo = http.Fetch(url)
        
        If Err.Number <> 0 Then
            Response.Write "<span class='error'>❌ ERRO no Fetch: " & Err.Description & "</span><br>"
        Else
            Response.Write "<span class='success'>✓ Fetch executado</span><br>"
        End If
        On Error GoTo 0
        
        Response.Write "</div>"
        
        ' Teste 3: Verificar o tipo de retorno
        Response.Write "<div class='debug-section'>"
        Response.Write "<h2>ETAPA 3: Analisar Retorno</h2>"
        
        Response.Write "<span class='label'>IsObject(todo):</span> <span class='value'>" & IsObject(todo) & "</span><br>"
        Response.Write "<span class='label'>TypeName(todo):</span> <span class='value'>" & TypeName(todo) & "</span><br>"
        Response.Write "<span class='label'>IsEmpty(todo):</span> <span class='value'>" & IsEmpty(todo) & "</span><br>"
        Response.Write "<span class='label'>IsNull(todo):</span> <span class='value'>" & IsNull(todo) & "</span><br>"
        
        Response.Write "</div>"
        
        ' Teste 4: Acessar propriedades
        Response.Write "<div class='debug-section'>"
        Response.Write "<h2>ETAPA 4: Acessar Propriedades</h2>"
        
        If IsObject(todo) Then
            On Error Resume Next
            
            Dim id, title, completed, userId
            
            id = todo("id")
            If Err.Number <> 0 Then
                Response.Write "<span class='error'>❌ Erro ao acessar 'id': " & Err.Description & "</span><br>"
            Else
                Response.Write "<span class='success'>✓ todo(""id""):</span> <span class='value'>" & id & "</span><br>"
            End If
            Err.Clear
            
            title = todo("title")
            If Err.Number <> 0 Then
                Response.Write "<span class='error'>❌ Erro ao acessar 'title': " & Err.Description & "</span><br>"
            Else
                Response.Write "<span class='success'>✓ todo(""title""):</span> <span class='value'>" & title & "</span><br>"
            End If
            Err.Clear
            
            completed = todo("completed")
            If Err.Number <> 0 Then
                Response.Write "<span class='error'>❌ Erro ao acessar 'completed': " & Err.Description & "</span><br>"
            Else
                Response.Write "<span class='success'>✓ todo(""completed""):</span> <span class='value'>" & completed & "</span><br>"
            End If
            Err.Clear
            
            userId = todo("userId")
            If Err.Number <> 0 Then
                Response.Write "<span class='error'>❌ Erro ao acessar 'userId': " & Err.Description & "</span><br>"
            Else
                Response.Write "<span class='success'>✓ todo(""userId""):</span> <span class='value'>" & userId & "</span><br>"
            End If
            
            On Error GoTo 0
        Else
            Response.Write "<span class='error'>❌ 'todo' NÃO é um objeto!</span><br>"
            Response.Write "<span class='label'>Tipo retornado:</span> <span class='value'>" & TypeName(todo) & "</span><br>"
            Response.Write "<span class='label'>Valor:</span> <span class='value'>" & todo & "</span><br>"
        End If
        
        Response.Write "</div>"
        
        ' Teste 5: Conteúdo bruto
        Response.Write "<div class='debug-section'>"
        Response.Write "<h2>ETAPA 5: Exibir Conteúdo Bruto</h2>"
        
        If IsObject(todo) Then
            Response.Write "<span class='label'>Objeto completo:</span> " & vbCrLf
            Response.Write "<pre class='value'>" & todo & "</pre>"
        Else
            Response.Write "<span class='label'>String/Valor:</span> " & vbCrLf
            Response.Write "<pre class='value'>" & todo & "</pre>"
        End If
        
        Response.Write "</div>"
        
        ' Teste 6: Verificar Env()
        Response.Write "<div class='debug-section'>"
        Response.Write "<h2>ETAPA 6: Verificar Env()</h2>"
        
        Dim apiKey
        apiKey = Env("API_KEY")
        Response.Write "<span class='success'>✓ Env(""API_KEY""):</span> <span class='value'>" & apiKey & "</span><br>"
        
        Response.Write "</div>"
    %>

</body>
</html>
