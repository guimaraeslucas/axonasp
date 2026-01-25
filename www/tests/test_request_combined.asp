<%@ LANGUAGE="VBSCRIPT" %>
<%
    ' Test both QueryString and Form together
    Response.Write("<h2>Combined Request Testing</h2>")
    
    Response.Write("<h3>QueryString Parameters:</h3>")
    If Request.QueryString.Count > 0 Then
        Response.Write("<ul>")
        Dim key
        For Each key In Request.QueryString
            Response.Write("<li>" & key & " = " & Request.QueryString(key) & "</li>")
        Next
        Response.Write("</ul>")
    Else
        Response.Write("<p>No QueryString parameters</p>")
    End If
    
    Response.Write("<h3>Form Parameters:</h3>")
    If Request.Form.Count > 0 Then
        Response.Write("<ul>")
        For Each key In Request.Form
            Response.Write("<li>" & key & " = " & Request.Form(key) & "</li>")
        Next
        Response.Write("</ul>")
    Else
        Response.Write("<p>No Form parameters</p>")
    End If
%>

<!DOCTYPE html>
<html>
<head>
    <title>Combined Request Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h2, h3 { color: #333; }
        form { background: #f5f5f5; padding: 15px; border-radius: 5px; margin: 20px 0; }
        input, textarea { display: block; margin: 10px 0; padding: 8px; width: 300px; }
        button { padding: 10px 20px; background: #28a745; color: white; border: none; border-radius: 3px; cursor: pointer; }
        button:hover { background: #218838; }
        ul { background: #f5f5f5; padding: 15px; border-radius: 5px; }
        li { margin: 5px 0; }
    </style>
</head>
<body>
    <h1>Combined QueryString and Form Testing</h1>
    
    <p><a href="/test_request_combined.asp?param1=value1&param2=value2">Test with QueryString only</a></p>
    
    <h3>Submit Form:</h3>
    <form method="POST" action="/test_request_combined.asp?param1=fromQueryString&param2=alsoQueryString">
        <input type="text" name="field1" placeholder="Field 1" required>
        <input type="text" name="field2" placeholder="Field 2" required>
        <input type="text" name="field3" placeholder="Field 3" required>
        <button type="submit">Submit with QueryString</button>
    </form>
</body>
</html>
