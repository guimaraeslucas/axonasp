<%@ Language="VBScript" %>
<html>
<head>
    <title>Template & Mail Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; }
        h1 { color: #0066cc; border-bottom: 3px solid #0066cc; padding-bottom: 10px; }
        h2 { color: #333; margin-top: 30px; }
        .success { color: #28a745; font-weight: bold; }
        .error { color: #dc3545; font-weight: bold; }
        .info { color: #666; padding: 10px; background: #f0f0f0; border-left: 4px solid #0066cc; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3 AxonASP - Template & Mail Libraries Test</h1>
        
        <h2>1. G3TEMPLATE Library</h2>
        <div class="info">
            <p><strong>Purpose:</strong> Template rendering using Go's html/template package</p>
            <p><strong>Usage:</strong> Server.CreateObject("G3TEMPLATE") or Server.CreateObject("TEMPLATE")</p>
        </div>
        <%
        Dim objTemplate, objMail
        
        ' Test G3TEMPLATE
        Set objTemplate = Server.CreateObject("G3TEMPLATE")
        If IsObject(objTemplate) Then
            Response.Write "<p class='success'>✓ G3TEMPLATE object created successfully</p>"
        Else
            Response.Write "<p class='error'>✗ Failed to create G3TEMPLATE object</p>"
        End If
        
        ' Test TEMPLATE alias
        Set objTemplate = Server.CreateObject("TEMPLATE")
        If IsObject(objTemplate) Then
            Response.Write "<p class='success'>✓ TEMPLATE (alias) object created successfully</p>"
        Else
            Response.Write "<p class='error'>✗ Failed to create TEMPLATE (alias) object</p>"
        End If
        %>
        
        <h2>2. G3MAIL Library</h2>
        <div class="info">
            <p><strong>Purpose:</strong> Send emails using SMTP</p>
            <p><strong>Usage:</strong> Server.CreateObject("G3MAIL") or Server.CreateObject("MAIL")</p>
            <p><strong>Methods:</strong></p>
            <ul>
                <li><strong>Send(host, port, username, password, from, to, subject, body, isHTML)</strong> - Send email with custom SMTP settings</li>
                <li><strong>SendStandard(to, subject, body, isHTML)</strong> - Send email using environment variables (SMTP_HOST, SMTP_PORT, etc.)</li>
            </ul>
        </div>
        <%
        ' Test G3MAIL
        Set objMail = Server.CreateObject("G3MAIL")
        If IsObject(objMail) Then
            Response.Write "<p class='success'>✓ G3MAIL object created successfully</p>"
        Else
            Response.Write "<p class='error'>✗ Failed to create G3MAIL object</p>"
        End If
        
        ' Test MAIL alias
        Set objMail = Server.CreateObject("MAIL")
        If IsObject(objMail) Then
            Response.Write "<p class='success'>✓ MAIL (alias) object created successfully</p>"
        Else
            Response.Write "<p class='error'>✗ Failed to create MAIL (alias) object</p>"
        End If
        %>
        
        <h2>3. Example Usage</h2>
        
        <h3>Template Example:</h3>
        <pre><%
Response.Write Server.HTMLEncode("Dim tpl") & vbCrLf
Response.Write Server.HTMLEncode("Set tpl = Server.CreateObject(""G3TEMPLATE"")") & vbCrLf
Response.Write Server.HTMLEncode("result = tpl.Render(""/templates/mytemplate.html"", data)") & vbCrLf
Response.Write Server.HTMLEncode("Response.Write result")
        %></pre>
        
        <h3>Mail Example (Custom SMTP):</h3>
        <pre><%
Response.Write Server.HTMLEncode("Dim mail") & vbCrLf
Response.Write Server.HTMLEncode("Set mail = Server.CreateObject(""G3MAIL"")") & vbCrLf
Response.Write Server.HTMLEncode("result = mail.Send(""smtp.gmail.com"", 587, ""user@gmail.com"", ""password"", _") & vbCrLf
Response.Write Server.HTMLEncode("    ""from@example.com"", ""to@example.com"", _") & vbCrLf
Response.Write Server.HTMLEncode("    ""Subject"", ""<h1>Hello</h1>"", True)")
        %></pre>
        
        <h3>Mail Example (Environment Variables):</h3>
        <pre><%
Response.Write Server.HTMLEncode("Dim mail") & vbCrLf
Response.Write Server.HTMLEncode("Set mail = Server.CreateObject(""G3MAIL"")") & vbCrLf
Response.Write Server.HTMLEncode("result = mail.SendStandard(""recipient@example.com"", ""Test Subject"", ""Email body"", False)")
        %></pre>
        
        <footer style="margin-top: 40px; padding-top: 20px; border-top: 2px solid #0066cc; text-align: center;">
            <p><strong>G3 AxonASP</strong> - Template & Mail Libraries Working</p>
        </footer>
    </div>
</body>
</html>
