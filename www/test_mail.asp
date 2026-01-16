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
                ' Configure SMTP settings (Use dummy data for safety/demo)
                host = "smtp.example.com"
                port = 587
                username = "user@example.com"
                password = "password"
                from = "sender@example.com"
                to = "recipient@example.com"
                subject = "Test Email from Go ASP (Manual)"
                body = "<h1>Hello!</h1><p>This is a test email sent manually.</p>"
                isHtml = true

                Response.Write("Target: " & to & "<br>")
                
                ' Call Mail.Send
                result = Mail.Send(host, port, username, password, from, to, subject, body, isHtml)

                If result = True Then
                    Response.Write("<span class='success'>Email enviado com sucesso (Mock)!</span>")
                Else
                    Response.Write("<span class='error'>Falha ao enviar: " & result & "</span>")
                End If
            %>
        </div>

        <div class="box">
            <h3>2. Envio via Env (Mail.SendStandard)</h3>
            <p>Tentando enviar usando variáveis de ambiente (SMTP_HOST, etc)...</p>
            <%
                toEnv = "recipient@example.com"
                subjectEnv = "Test Email from Go ASP (Env)"
                bodyEnv = "This uses the .env configuration."
                
                resultEnv = mail.SendStandard(toEnv, subjectEnv, bodyEnv, True)

                If resultEnv = True Then
                     Response.Write("<span class='success'>Email enviado com sucesso!</span>")
                Else
                     Response.Write("<span class='error'>Falha ao enviar: " & resultEnv & "</span>")
                End If
            %>
        </div>
        
        <p><a href="default.asp">Voltar para Home</a></p>
    </div>
</body>
</html>