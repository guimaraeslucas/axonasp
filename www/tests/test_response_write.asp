<%
' Test Response.Write Implementation
' Demonstra múltiplos tipos de dados e conversões

Response.Write("<h1>Response.Write Test</h1>")
Response.Write("<hr>")

' Test 1: String output
Response.Write("<h2>1. String Output</h2>")
Response.Write("<p>Hello World!</p>")

' Test 2: Número inteiro
Response.Write("<h2>2. Integer Output</h2>")
Dim num
num = 42
Response.Write("<p>The answer is: " & num & "</p>")

' Test 3: Número decimal
Response.Write("<h2>3. Float Output</h2>")
Dim price
price = 19.99
Response.Write("<p>Price: $" & price & "</p>")

' Test 4: Boolean
Response.Write("<h2>4. Boolean Output</h2>")
Dim isActive
isActive = True
Response.Write("<p>Active: " & isActive & "</p>")

' Test 5: Múltiplos argumentos
Response.Write("<h2>5. Multiple Arguments</h2>")
Response.Write("<p>")
Response.Write("Name: ")
Response.Write("John")
Response.Write(" | Age: ")
Response.Write(30)
Response.Write("</p>")

' Test 6: Complex HTML
Response.Write("<h2>6. Complex HTML</h2>")
Response.Write("<table border='1'>")
Response.Write("<tr><th>Item</th><th>Price</th></tr>")
Response.Write("<tr><td>Apple</td><td>$1.50</td></tr>")
Response.Write("<tr><td>Orange</td><td>$2.00</td></tr>")
Response.Write("</table>")

' Test 7: Empty and null values
Response.Write("<h2>7. Empty Values</h2>")
Response.Write("<p>Empty string: [" & "" & "]</p>")

Response.Write("<hr>")
Response.Write("<p><strong>All Response.Write tests completed!</strong></p>")
%>
