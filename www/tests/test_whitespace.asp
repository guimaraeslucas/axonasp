<!DOCTYPE html>
<html>
<head>
    <title>Parser Whitespace Test</title>
</head>
<body>
    <h1>Teste de Espaços Antes de &lt;%</h1>
    
    <div>
        <h2>Bloco 1: Sem espaços</h2>
        <%
        Dim x
        x = 1
        Response.Write "X = " & x & "<br>"
        %>
    </div>

    <div>
        <h2>Bloco 2: Com espaços antes de &lt;%</h2>
            <%
            Dim y
            y = 2
            Response.Write "Y = " & y & "<br>"
            %>
    </div>

    <div>
        <h2>Bloco 3: Com muitos espaços e quebras</h2>
        
            
                <%
                Dim z
                z = 3
                Response.Write "Z = " & z & "<br>"
                %>
    </div>

    <div>
        <h2>Bloco 4: Select Case</h2>
        <%
        Dim day
        day = 1
        Select Case day
            Case 1
                Response.Write "Monday<br>"
            Case 2
                Response.Write "Tuesday<br>"
            Case Else
                Response.Write "Other<br>"
        End Select
        %>
    </div>

    <div>
        <h2>Bloco 5: Do Loop válido</h2>
        <%
        Dim counter
        counter = 0
        Do Until counter >= 3
            Response.Write "Counter = " & counter & "<br>"
            counter = counter + 1
        Loop
        %>
    </div>

    <p>✓ Teste concluído com sucesso!</p>
</body>
</html>
