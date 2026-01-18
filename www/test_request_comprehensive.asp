<%@ LANGUAGE="VBSCRIPT" %>
<%
    Response.Write("<h2>Comprehensive Request Object Testing</h2>")
    
    ' Test 1: QueryString with parameters
    Response.Write("<h3>Test 1: QueryString Access</h3>")
    Response.Write("<p>Direct access to 'name' parameter: " & Request.QueryString("name") & "</p>")
    Response.Write("<p>Direct access to 'age' parameter: " & Request.QueryString("age") & "</p>")
    
    ' Test 2: QueryString iteration
    Response.Write("<h3>Test 2: QueryString Iteration</h3>")
    Response.Write("<p>All QueryString parameters:</p>")
    Response.Write("<ul>")
    Dim param
    For Each param In Request.QueryString
        Response.Write("<li>" & param & " = " & Request.QueryString(param) & "</li>")
    Next
    Response.Write("</ul>")
    
    ' Test 3: Check if parameter exists
    Response.Write("<h3>Test 3: Parameter Existence Check</h3>")
    If Request.QueryString("name") <> "" Then
        Response.Write("<p>Name parameter exists: " & Request.QueryString("name") & "</p>")
    Else
        Response.Write("<p>Name parameter does not exist or is empty</p>")
    End If
    
    ' Test 4: ServerVariables
    Response.Write("<h3>Test 4: ServerVariables</h3>")
    Response.Write("<ul>")
    Response.Write("<li>REQUEST_METHOD: " & Request.ServerVariables("REQUEST_METHOD") & "</li>")
    Response.Write("<li>REQUEST_PATH: " & Request.ServerVariables("REQUEST_PATH") & "</li>")
    Response.Write("<li>QUERY_STRING: " & Request.ServerVariables("QUERY_STRING") & "</li>")
    Response.Write("<li>REMOTE_ADDR: " & Request.ServerVariables("REMOTE_ADDR") & "</li>")
    Response.Write("</ul>")
%>

<!DOCTYPE html>
<html>
<head>
    <title>Comprehensive Request Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h2, h3 { color: #333; }
        ul { background: #f5f5f5; padding: 15px; border-radius: 5px; margin: 10px 0; }
        li { margin: 5px 0; }
        a { color: #007bff; text-decoration: none; }
        a:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <h1>Comprehensive Request Object Test</h1>
    
    <p>Test links with different parameters:</p>
    <ul>
        <li><a href="/test_request_comprehensive.asp">No parameters</a></li>
        <li><a href="/test_request_comprehensive.asp?name=John&age=25">With name and age</a></li>
        <li><a href="/test_request_comprehensive.asp?name=Jane&age=30&city=NewYork">With name, age, and city</a></li>
        <li><a href="/test_request_comprehensive.asp?single=value">With single parameter</a></li>
    </ul>
</body>
</html>
