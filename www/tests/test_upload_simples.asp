<%
Option Explicit

Dim isPost, action, uploader, resultObj, resultHtml

isPost = (Request.ServerVariables("REQUEST_METHOD") = "POST")
action = Trim(Request.Form("action"))

resultHtml = ""

' DEBUG: Write method and action to console
Response.Write "<!-- DEBUG: isPost=" & isPost & ", action=[" & action & "] -->" & vbCrLf

' Test upload
If isPost And action = "simple" Then
    resultHtml = "<!-- DEBUG: Entering simple upload handler -->" & vbCrLf
    
    Set uploader = Server.CreateObject("G3FileUploader")
    uploader.SetProperty "debugmode", True
    uploader.BlockExtensions "exe,dll,bat,cmd"
    uploader.SetProperty "maxfilesize", 50485760
    
    Set resultObj = uploader.Process("file1", "/uploads")
    
    If resultObj("IsSuccess") = True Then
        resultHtml = resultHtml & "<div style='color:green;padding:10px;background:#ffffcc;'><strong>✓ SUCESSO!</strong><br>" & _
                    "Arquivo: " & resultObj("OriginalFileName") & "<br>" & _
                    "Novo nome: " & resultObj("NewFileName") & "<br>" & _
                    "Tamanho: " & resultObj("Size") & " bytes</div>"
    Else
        resultHtml = resultHtml & "<div style='color:red;padding:10px;background:#ffcccc;'><strong>✗ ERRO!</strong><br>" & _
                    "Mensagem: " & resultObj("ErrorMessage") & "</div>"
    End If
End If
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>G3FileUploader Test</title>
</head>
<body>
    <h1>G3FileUploader - Teste</h1>
    
    <div style="margin:20px;padding:20px;border:1px solid #ccc;">
        <h2>Upload Simples</h2>
        <% 
        If resultHtml <> "" Then
            Response.Write resultHtml
        End If
        %>
        
        <form method="POST" enctype="multipart/form-data" style="margin-top:20px;">
            <input type="hidden" name="action" value="simple">
            <input type="file" name="file1" required style="margin:10px 0;">
            <input type="submit" value="Enviar Arquivo" style="padding:10px 20px;">
        </form>
    </div>
    
    <div style="margin:20px;padding:20px;border:1px solid #ccc;background:#f5f5f5;">
        <h3>Status do Servidor</h3>
        <p>Host: <%= Request.ServerVariables("SERVER_NAME") %></p>
        <p>Método: <%= Request.ServerVariables("REQUEST_METHOD") %></p>
        <p>Action recebida: <%= action %></p>
    </div>
</body>
</html>
