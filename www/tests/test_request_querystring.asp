<%@ LANGUAGE="VBSCRIPT" %>
<%
    ' Test QueryString Collection
    Response.Write("<h2>Test QueryString Collection</h2>")
    
    ' Access individual QueryString parameter
    Response.Write("<p><strong>Single access:</strong></p>")
    Response.Write("<p>msg parameter: " & Request.QueryString("msg") & "</p>")
    Response.Write("<p>action parameter: " & Request.QueryString("action") & "</p>")
    Response.Write("<p>non-existent parameter: " & Request.QueryString("nonexistent") & "</p>")
    
    ' Test For Each iteration over QueryString
    Response.Write("<p><strong>For Each iteration:</strong></p>")
    Response.Write("<ul>")
    
    Dim key
    For Each key In Request.QueryString
        Response.Write("<li>" & key & " = " & Request.QueryString(key) & "</li>")
    Next
    
    Response.Write("</ul>")
%>
<!DOCTYPE html>
<html>
<head>
    <title>QueryString Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h2 { color: #333; }
        ul { background: #f5f5f5; padding: 15px; border-radius: 5px; }
        li { margin: 5px 0; }
    </style>
</head>
<body>
    <h1>Request.QueryString Testing</h1>
    <p><a href="/test_request_querystring.asp?msg=hello&action=test">Test with parameters</a></p>
    <p><a href="/test_request_querystring.asp">Test without parameters</a></p>
</body>
</html>
