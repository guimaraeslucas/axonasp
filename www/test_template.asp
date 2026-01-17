<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Template Library Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .box { border-left: 4px solid #667eea; padding: 15px; margin-bottom: 15px; background: #f9f9f9; border-radius: 4px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .render-area { border: 1px dashed #ccc; padding: 15px; background: #fff; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Template Engine Test</h1>
        <div class="intro">
            <p>Demonstrates rendering external <code>.html</code> files using data processed in ASP (VBScript).</p>
        </div>

        <div class="box">
            <h3>Preparando Dados</h3>
            <%
                Dim json
                Set json = Server.CreateObject("G3JSON")
                
                ' Create Data Object
                Set data = json.NewObject()
                data("Title") = "Go ASP Template Demo"
                data("Header") = "Dynamic Page via Go Templates"

                ' Nested Object
                Set user = json.NewObject()
                user("Name") = "Alice"
                user("Role") = "Administrator"
                data("User") = user

                ' Array
                Set items = json.NewArray()
                items(0) = "Apple"
                items(1) = "Banana"
                items(2) = "Cherry"
                data("Items") = items
                
                Response.Write("Dados JSON preparados com sucesso.<br>")
            %>
        </div>

        <div class="box">
            <h3>Renderização</h3>
            <p>Abaixo está o resultado de <code>Template.Render("test_template.html", data)</code>:</p>
            <div class="render-area">
            <%
                Dim tmpl
                Set tmpl = Server.CreateObject("G3TEMPLATE")
                
                ' Render
                output = tmpl.Render("test_template.html", data)
                Response.Write(output)
            %>
            </div>
        </div>

        <p><a href="default.asp">Voltar para Home</a></p>
    </div>
</body>
</html>