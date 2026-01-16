<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Test: Simple pre() Function</title>
</head>
<body>
    <h1>Test Pre Function</h1>
    
    <%
    Function pre(value)
        Dim pre
        pre = Replace(value, vbTab, " ")
    End Function
    %>
    
    <p>Calling pre() via output tag:</p>
    <p><%= pre("test") %></p>
    
</body>
</html>
