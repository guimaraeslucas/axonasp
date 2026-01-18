<%@ LANGUAGE="VBSCRIPT" %>
<%
    ' Test Form Collection - demonstrates Request.Form functionality
    Response.Write("<h2>Test Form Collection</h2>")
    
    ' Always show form data if present
    Response.Write("<h3>Form Data Received:</h3>")
    Response.Write("<ul>")
    
    Dim key
    Dim hasData
    hasData = False
    
    For Each key In Request.Form
        hasData = True
        Response.Write("<li>" & key & " = " & Request.Form(key) & "</li>")
    Next
    
    Response.Write("</ul>")
    
    ' Show direct access examples
    Response.Write("<h3>Direct Access Examples:</h3>")
    Response.Write("<p>username: " & Request.Form("username") & "</p>")
    Response.Write("<p>password: " & Request.Form("password") & "</p>")
    Response.Write("<p>email: " & Request.Form("email") & "</p>")
%>

<!DOCTYPE html>
<html>
<head>
    <title>Form Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h2 { color: #333; }
        form { background: #f5f5f5; padding: 15px; border-radius: 5px; margin: 20px 0; }
        input, textarea { display: block; margin: 10px 0; padding: 8px; width: 300px; }
        button { padding: 10px 20px; background: #007bff; color: white; border: none; border-radius: 3px; cursor: pointer; }
        button:hover { background: #0056b3; }
        ul { background: #f5f5f5; padding: 15px; border-radius: 5px; }
        li { margin: 5px 0; }
    </style>
</head>
<body>
    <h1>Request.Form Testing</h1>
    
    <form method="POST" action="/test_request_form.asp">
        <input type="text" name="username" placeholder="Username" required>
        <input type="password" name="password" placeholder="Password" required>
        <input type="email" name="email" placeholder="Email" required>
        <input type="hidden" name="submitted" value="1">
        <button type="submit">Submit Form</button>
    </form>
</body>
</html>
