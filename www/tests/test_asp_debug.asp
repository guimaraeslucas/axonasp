<!DOCTYPE html>
<html>
<head>
    <title>ASP Debug Test</title>
</head>
<body>
    <h1>Testing ASP Execution</h1>
    <p>This should show HTML always.</p>
    <%
    Response.Write("This is from ASP code!")
    Response.Write("<br>")
    Response.Write("If you see this, ASP is working!")
    %>
    <p>And this HTML should appear after ASP.</p>
</body>
</html>
