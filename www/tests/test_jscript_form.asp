<%@Language="JScript"%>
<%
    // Captura o método da requisição para identificar se houve envio (POST)
    var metodo = String(Request.ServerVariables("REQUEST_METHOD"));
    var nomeEnviado = "";

    if (metodo === "POST") {
        // Em JScript, é necessário converter o objeto Request.Form explicitamente para String
        nomeEnviado = String(Request.Form("nome"));
    }
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Exemplo ASP Clássico com JScript</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 40px;
        }
        .resultado {
            padding: 15px;
            border-left: 5px solid #4CAF50;
            background-color: #f1f8e9;
            margin-bottom: 20px;
        }
        .formulario {
            margin-top: 20px;
        }
        input[type="text"] {
            padding: 8px;
            margin-top: 5px;
            width: 250px;
        }
        button {
            padding: 8px 15px;
            margin-top: 10px;
            cursor: pointer;
        }
    </style>
</head>
<body>

    <h2>Formulário de Saudação</h2>

    <% 
    // Bloco condicional JScript para renderizar a resposta ou o formulário
    if (metodo === "POST" && nomeEnviado !== "") { 
    %>
        <div class="resultado">
            <strong>Sucesso!</strong><br>
            Olá, <strong><%= nomeEnviado %></strong>! Seu nome foi recebido pelo servidor.
        </div>
        <a href="?">Voltar e testar novamente</a>
    <% 
    } else { 
    %>
        <div class="formulario">
            <form method="POST" action="">
                <label for="nome">Qual é o seu nome?</label><br>
                <input type="text" id="nome" name="nome" autocomplete="off" required>
                <br>
                <button type="submit">Enviar</button>
            </form>
        </div>
    <% 
    } 
    %>

</body>
</html>