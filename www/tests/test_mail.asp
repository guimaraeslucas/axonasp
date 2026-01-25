<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Mail Library Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .box { border-left: 4px solid #667eea; padding: 15px; margin-bottom: 15px; background: #f9f9f9; border-radius: 4px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .success { color: green; font-weight: bold; }
        .error { color: red; font-weight: bold; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Mail Library Test</h1>
        <div class="intro">
            <p>Tests email sending capabilities using the integrated gomail.v2 library.</p>
            <p>Check the <code>.env</code> file for "Standard" mode configuration.</p>
        </div>
        
        <div class="box">
            <h3>1. Envio Manual (Mail.Send)</h3>
            <p>Tentando enviar usando dados hardcoded (Simulação)...</p>
            <%
                Dim mail, result
                Set mail = Server.CreateObject("G3MAIL")
                
                If IsObject(mail) Then
                    Response.Write "<span class='success'>✓ Objeto G3MAIL criado com sucesso</span><br>"
                    Response.Write "<p><strong>Método Send:</strong> mail.Send(host, port, username, password, from, to, subject, body, isHTML)</p>"
                    Response.Write "<p><em>Nota: Para testar o envio real, configure os parâmetros SMTP corretamente.</em></p>"
                Else
                    Response.Write "<span class='error'>✗ Falha ao criar objeto G3MAIL</span>"
                End If
            %>
        </div>

        <div class="box">
            <h3>2. Envio via Env (Mail.SendStandard)</h3>
            <p>Tentando enviar usando variáveis de ambiente (SMTP_HOST, etc)...</p>
            <%
                If IsObject(mail) Then
                    Response.Write "<span class='success'>✓ Objeto G3MAIL disponível</span><br>"
                    Response.Write "<p><strong>Método SendStandard:</strong> mail.SendStandard(to, subject, body, isHTML)</p>"
                    Response.Write "<p><em>Nota: Configure as variáveis SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM no arquivo .env</em></p>"
                Else
                    Response.Write "<span class='error'>✗ Objeto Mail não disponível</span>"
                End If
            %>
        </div>
        
        <p><a href="default.asp">Voltar para Home</a></p>
    </div>
</body>
</html>