<%@ Language=VBScript %>
<html>
<head><title>Short Write Test</title></head>
<body>
    <h3>Testing Short Write Syntax</h3>
    <%
    Dim name
    name = "World"
    Dim num
    num = 42
    %>
    <p>Hello, <%= name %>!</p>
    <p>The answer is <%= num %>.</p>
    <p>Math: 5 + 5 = <%= 5 + 5 %></p>
    <p>Time: <%= Now %></p>
    <p>Concatenation: <%= "Hello " & "ASP" %></p>
</body>
</html>
