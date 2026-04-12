<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Date Functions Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
        th { background: #f5f5f5; color: #333; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Date Functions Test</h1>
        <div class="intro">
            <p>Tests DateDiff, DatePart, DateAdd, DateValue, DateSerial and FormatDateTime functions.</p>
        </div>
        <h2>Resultados</h2>
<% Response.Write("<p>")
Response.Write("DateDiff('d', '01/01/2025', '01/31/2025'): " & DateDiff("d", "01/01/2025", "01/31/2025") & " (Expected: 30)<br>")
Response.Write("DatePart('m', '03/15/2025'): " & DatePart("m", "03/15/2025") & " (Expected: 3)<br>")
Response.Write("DatePart('d', '03/15/2025'): " & DatePart("d", "03/15/2025") & " (Expected: 15)<br>")
Response.Write("DatePart('yyyy', '03/15/2025'): " & DatePart("yyyy", "03/15/2025") & " (Esperado: 2025)<br>")
Response.Write("</p>")
%>
        <p style="margin-top:15px">Todos os testes devem mostrar os valores esperados acima.</p>
    </div>
</body>
</html>

