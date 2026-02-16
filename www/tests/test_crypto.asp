<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Cryptography Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .code-block { background: #f4f4f4; padding: 10px; margin: 10px 0; border-radius: 4px; word-break: break-all; }
        code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; }
        .success { color: #28a745; }
        .error { color: #dc3545; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Cryptography & UUID Test</h1>
        <div class="intro">
            <p>Tests UUID generation, BCrypt password hashing and verification.</p>
        </div>

    <%
        Dim crypt
        Set crypt = Server.CreateObject("G3CRYPTO")

        ' 1. Teste de UUID
        Response.Write "<h3>1. Gerando UUIDs</h3>"
        Dim id1, id2
        id1 = crypt.UUID()
        id2 = crypt.UUID()
        
        Response.Write "UUID 1: <code>" & id1 & "</code><br>"
        Response.Write "UUID 2: <code>" & id2 & "</code><br>"
        
        If id1 <> id2 Then
            Response.Write "<span style='color:green'>Sucesso: UUIDs são únicos.</span><br>"
        End If

        ' 2. Teste de Hash de Senha
        Response.Write "<h3>2. Hashing de Senha (BCrypt)</h3>"
        Dim senha, hash
        senha = "minhaSenhaSuperSecreta123"
        
        ' Gera o hash
        hash = crypt.HashPassword(senha)
        Response.Write "Senha: " & senha & "<br>"
        Response.Write "Hash Gerado: <div style='background:#eee; padding:5px; word-break:break-all;'>" & hash & "</div>"

        ' 3. Verificação Correta
        Response.Write "<h3>3. Verificando Senha</h3>"
        
        If crypt.VerifyPassword(senha, hash) Then
            Response.Write "<span style='color:green'>[OK] A senha correta foi validada.</span><br>"
        Else
            Response.Write "<span style='color:red'>[ERRO] A senha correta falhou.</span><br>"
        End If

        ' 4. Verificação Incorreta
        If crypt.VerifyPassword("senhaErrada", hash) Then
             Response.Write "<span style='color:red'>[ERRO] Senha errada foi aceita!</span><br>"
        Else
             Response.Write "<span style='color:green'>[OK] Senha errada foi rejeitada corretamente.</span><br>"
        End If
    %>
    </div>
</body>
</html>