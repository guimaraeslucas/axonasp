<%@ Language=VBScript %>
<html>
<head>
    <title>GoLang ASP Interpreter</title>
    <style>
        body { font-family: sans-serif; padding: 20px; }
        .box { border: 1px solid #ccc; padding: 10px; margin-bottom: 10px; background: #f9f9f9; }
        h3 { margin-top: 0; }
    </style>
</head>
<body>
<%
Response.Write("Hello World!")
%>
<%="<p style='color:#0000ff'>This text is styled.</p>"%>


    <div class="box">
        <h3>1. Response.Write & Variables</h3>
        <%
            Dim name
            name = "Visitor"
            'Response.Write("Hello, " & name)
            Response.Write "<br>Current Server Time: " & Now() 
            ' Note: Now() calls internal Go time in a real engine, 
            ' here it might just print string "Now()" if not implemented, 
            ' or we relies on the evaluator returning it literally.
            ' Let's test a simple string concat.
        %>
    </div>

    <div class="box">
        <h3>2. Logic (If/Else)</h3>
        <%
            Dim score
            score = 85
            Response.Write "Score is: " & score & "<br>"
            
            If score > 50 Then
                Response.Write "<b>Pass!</b> The score is greater than 50."
            Else
                Response.Write "<b>Fail!</b>"
            End If
            
            Response.Write "<br>"
            
            If score < 10 Then
                Response.Write "Very Low"
            Else
                Response.Write "Not Low"
            End If
        %>
    </div>

    <div class="box">
        <h3>3. Loops (For ... Next)</h3>
        <ul>
        <%
            Dim i
            For i = 1 To 5
                Response.Write "<li>Item number " & i & "</li>"
            Next
        %>
        </ul>
    </div>

    <div class="box">
        <h3>4. Subroutines (Call)</h3>
        <%
            Sub MySub
                Response.Write "I am inside a Subroutine!<br>"
            End Sub

            Response.Write "Calling Sub...<br>"
            Call MySub
            Response.Write "Back from Sub."
        %>
    </div>

    <div class="box">
        <h3>5. Session & Application</h3>
        <%
            ' Counter in Session
            Dim count
            count = Session("hits")
            If count = "" Then count = 0
            count = count + 1
            Session("hits") = count
            
            Response.Write "Your Session Hits: " & count & "<br>"
            Response.Write "Session ID: " & Session.SessionID & "<br>"

            ' Application Global Counter
            Dim total
            total = Application("total_hits")
            If total = "" Then total = 0
            total = total + 1
            Application("total_hits") = total
            
            Response.Write "Total Server Hits (All Users): " & total
        %>
    </div>
</body>
</html>
