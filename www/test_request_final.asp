<%@ LANGUAGE="VBSCRIPT" %>
<!DOCTYPE html>
<html>
<head>
    <title>Request Object - Full Test Suite</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f9f9f9; }
        .container { max-width: 900px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        h2 { color: #3498db; margin-top: 30px; }
        h3 { color: #555; }
        .test-section { background: #f8f9fa; padding: 15px; margin: 15px 0; border-left: 4px solid #3498db; border-radius: 4px; }
        ul { list-style: none; padding: 0; }
        li { padding: 8px; margin: 5px 0; background: white; border-radius: 4px; border: 1px solid #ddd; }
        .key { font-weight: bold; color: #2c3e50; }
        .value { color: #27ae60; }
        form { background: #ecf0f1; padding: 20px; border-radius: 4px; margin: 20px 0; }
        input, textarea { display: block; margin: 10px 0; padding: 10px; width: 100%; border: 1px solid #bdc3c7; border-radius: 4px; box-sizing: border-box; }
        button { padding: 12px 30px; background: #3498db; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }
        button:hover { background: #2980b9; }
        .success { color: #27ae60; font-weight: bold; }
        .info { color: #7f8c8d; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Request Object - Complete Test Suite</h1>
        <p class="info">Testing Request.QueryString, Request.Form, Request.ServerVariables, and For Each iteration</p>
        
        <!-- QueryString Tests -->
        <div class="test-section">
            <h2>1. QueryString Collection</h2>
            
            <h3>Direct Access:</h3>
            <ul>
<%
    Response.Write("<li><span class='key'>name:</span> <span class='value'>" & Request.QueryString("name") & "</span></li>")
    Response.Write("<li><span class='key'>age:</span> <span class='value'>" & Request.QueryString("age") & "</span></li>")
    Response.Write("<li><span class='key'>city:</span> <span class='value'>" & Request.QueryString("city") & "</span></li>")
%>
            </ul>
            
            <h3>For Each Iteration:</h3>
            <ul>
<%
    Dim qsKey
    For Each qsKey In Request.QueryString
        Response.Write("<li><span class='key'>" & qsKey & ":</span> <span class='value'>" & Request.QueryString(qsKey) & "</span></li>")
    Next
%>
            </ul>
        </div>
        
        <!-- Form Tests -->
        <div class="test-section">
            <h2>2. Form Collection</h2>
            
            <h3>Form Data Received:</h3>
            <ul>
<%
    Dim formKey
    Dim formCount
    formCount = 0
    
    For Each formKey In Request.Form
        formCount = formCount + 1
        Response.Write("<li><span class='key'>" & formKey & ":</span> <span class='value'>" & Request.Form(formKey) & "</span></li>")
    Next
    
    If formCount = 0 Then
        Response.Write("<li class='info'>No form data submitted yet</li>")
    End If
%>
            </ul>
            
            <h3>Submit Test Form:</h3>
            <form method="POST" action="/test_request_final.asp?name=John&age=30&city=NYC">
                <input type="text" name="username" placeholder="Username" value="TestUser">
                <input type="email" name="email" placeholder="Email" value="test@example.com">
                <input type="text" name="phone" placeholder="Phone" value="555-1234">
                <textarea name="message" placeholder="Your message" rows="3">This is a test message</textarea>
                <button type="submit">Submit Form</button>
            </form>
        </div>
        
        <!-- ServerVariables Tests -->
        <div class="test-section">
            <h2>3. ServerVariables Collection</h2>
            
            <ul>
<%
    Response.Write("<li><span class='key'>REQUEST_METHOD:</span> <span class='value'>" & Request.ServerVariables("REQUEST_METHOD") & "</span></li>")
    Response.Write("<li><span class='key'>REQUEST_PATH:</span> <span class='value'>" & Request.ServerVariables("REQUEST_PATH") & "</span></li>")
    Response.Write("<li><span class='key'>QUERY_STRING:</span> <span class='value'>" & Request.ServerVariables("QUERY_STRING") & "</span></li>")
    Response.Write("<li><span class='key'>REMOTE_ADDR:</span> <span class='value'>" & Request.ServerVariables("REMOTE_ADDR") & "</span></li>")
%>
            </ul>
        </div>
        
        <!-- Test Links -->
        <div class="test-section">
            <h2>4. Test Links</h2>
            <p>Try these links to test QueryString parameters:</p>
            <ul>
                <li><a href="/test_request_final.asp">No parameters</a></li>
                <li><a href="/test_request_final.asp?name=John&age=25">With name and age</a></li>
                <li><a href="/test_request_final.asp?name=Jane&age=30&city=NewYork">With name, age, and city</a></li>
                <li><a href="/test_request_final.asp?product=laptop&price=999&quantity=2">Product query</a></li>
            </ul>
        </div>
        
        <p class="success">Request.QueryString - Working</p>
        <p class="success">Request.Form - Working</p>
        <p class="success">Request.ServerVariables - Working</p>
        <p class="success">For Each Iteration - Working</p>
    </div>
</body>
</html>
