<!DOCTYPE html>
<html>
<head>
    <title>Session Test</title>
    <style>
        body { font-family: Arial; padding: 20px; }
        .box { border: 1px solid #ddd; padding: 15px; margin: 10px 0; background: #f9f9f9; }
    </style>
</head>
<body>
    <h1>G3pix AxonASP - Session Test</h1>
    
    <div class="box">
        <h3>1. Basic Session Variables</h3>
        <%
            ' Test Session variable assignment and retrieval
            Session("user_name") = "John Doe"
            Dim saved_name
            saved_name = Session("user_name")
            Response.Write "Stored in Session: " & saved_name & "<br>"
        %>
    </div>
    
    <div class="box">
        <h3>2. Session Counter</h3>
        <%
            ' Counter using Session
            Dim count
            count = Session("visit_count")
            
            ' Check if variable exists and is empty/nil
            If count = "" Or IsNull(count) Then
                count = 0
            End If
            
            count = count + 1
            Session("visit_count") = count
            
            Response.Write "Visit Count: " & count & "<br>"
        %>
    </div>
    
    <div class="box">
        <h3>3. Session ID</h3>
        <%
            ' Display session ID
            Response.Write "Session ID: " & Session.SessionID & "<br>"
        %>
    </div>
    
    <div class="box">
        <h3>4. Multiple Session Values</h3>
        <%
            ' Store and retrieve multiple values
            Session("email") = "user@example.com"
            Session("age") = 25
            Session("active") = True
            
            Response.Write "Email: " & Session("email") & "<br>"
            Response.Write "Age: " & Session("age") & "<br>"
            Response.Write "Active: " & Session("active") & "<br>"
        %>
    </div>
    
    <p>
        <a href="test_session.asp">Refresh page</a> to see if session values persist across requests.
    </p>
    
</body>
</html>
