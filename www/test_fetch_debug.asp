<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Debug Fetch</title>
    <style>
        body { font-family: Arial; padding: 20px; background: #f5f5f5; }
        .container { max-width: 800px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; }
        pre { background: #f4f4f4; padding: 15px; overflow-x: auto; }
        .info { color: #0066cc; }
        .success { color: green; }
        .error { color: red; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Debug Fetch</h1>
        
        <h2>Teste 1: Verificar tipo de retorno</h2>
        <%
            Dim http, resultado, resultadoType
            Set http = Server.CreateObject("G3HTTP")
            Set resultado = http.Fetch("https://jsonplaceholder.typicode.com/todos/1")
            
            resultadoType = TypeName(resultado)
            
            Response.Write "<p><strong>Tipo retornado:</strong> <span class='info'>" & resultadoType & "</span></p>"
            Response.Write "<p><strong>IsObject:</strong> " & IsObject(resultado) & "</p>"
            Response.Write "<p><strong>IsEmpty:</strong> " & IsEmpty(resultado) & "</p>"
            Response.Write "<p><strong>IsNull:</strong> " & IsNull(resultado) & "</p>"
            
            If resultado <> "" Then
                Response.Write "<p class='success'><strong>✓ Resultado não está vazio</strong></p>"
            Else
                Response.Write "<p class='error'><strong>✗ Resultado está vazio</strong></p>"
            End If
        %>
        
        <h2>Teste 2: Ver conteúdo bruto</h2>
        <%
            Response.Write "<pre>"
            Response.Write "Conteúdo retornado: " & vbCrLf
            Response.Write resultado
            Response.Write "</pre>"
        %>
        
        <h2>Teste 3: Testar acesso a propriedades</h2>
        <%
            If IsObject(resultado) Then
                Response.Write "<p class='success'>É um objeto!</p>"
                Response.Write "<p><strong>ID:</strong> " & resultado("id") & "</p>"
                Response.Write "<p><strong>Título:</strong> " & resultado("title") & "</p>"
                Response.Write "<p><strong>Completed:</strong> " & resultado("completed") & "</p>"
            Else
                Response.Write "<p class='error'>NÃO é um objeto</p>"
                
                ' Tentar converter manualmente
                Response.Write "<p>Tentando fazer parse manual...</p>"
                Dim json
                Set json = Server.CreateObject("G3JSON")
                Dim parsed
                Set parsed = json.Parse(CStr(resultado))
                
                If IsObject(parsed) Then
                    Response.Write "<p class='success'>✓ Parse bem-sucedido!</p>"
                    Response.Write "<p><strong>ID:</strong> " & parsed("id") & "</p>"
                    Response.Write "<p><strong>Título:</strong> " & parsed("title") & "</p>"
                Else
                    Response.Write "<p class='error'>✗ Parse falhou</p>"
                End If
            End If
        %>
    </div>
</body>
</html>
