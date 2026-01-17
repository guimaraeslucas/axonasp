
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Basic Syntax Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .box { border-left: 4px solid #667eea; padding: 15px; margin-bottom: 15px; background: #f9f9f9; border-radius: 4px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        pre { background: #f4f4f4; padding: 10px; border-radius: 4px; overflow-x: auto; }
        code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; }
        ul { margin-left: 20px; }
        li { margin-bottom: 5px; }
        form { margin: 15px 0; }
        input, button { padding: 8px 12px; margin: 5px 5px 5px 0; border: 1px solid #ddd; border-radius: 4px; }
        button { background: #667eea; color: #fff; cursor: pointer; }
        button:hover { background: #764ba2; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Basic Syntax & Logic Test</h1>
        <div class="intro">
            <p>Tests variables, control flow (If/Else), loops, subroutines, Session, Arrays and string operations.</p>
        </div>

    <div class="box">
        <h3>1. Response.Write & Variables</h3>
        <%
            Dim name
            name = "Visitor"
            Response.Write("Hello, " & name)
            Response.Write "<br>Current Server Time: " & Now() 
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

            Dim x 
            If (IsEmpty(x)) Then
                x=10
                Response.Write "<br>x was empty, set to 10<br>"
            End If
            Response.Write(x)
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
        <ul>
            <%
            d=Day()
            Select Case d
            Case 1
            response.write("Sleepy Sunday")
            Case 2
            response.write("Monday again!")
            Case 3
            response.write("Just Tuesday!")
            Case 4
            response.write("Wednesday!")
            Case 5
            response.write("Thursday...")
            Case 6
            response.write("Finally Friday!")
            Case else
            response.write("Super Saturday!!!!")
            End Select

            response.write("<br>Dia:" & d)

             %>
        </ul>
        <ul>
            <%
            Dim i_loop
            i_loop = 0
            Do Until i_loop > 2
            i_loop = i_loop + 1
            If i_loop > 2 Then Exit Do
            Loop
             %>
        </ul>
        </ul>
    </div>

    <div class="box">
        <h3>4. Subroutines (Call & Math)</h3>
        <%
            Sub MySub(n1,n2)
                Response.Write "I am inside a Subroutine!<br>"
                Response.Write ("Result of " & n1 & " * " & n2 & " is: " & (n1*n2) & "<br>")
            End Sub

            Response.Write "Calling Sub with 2 and 3...<br>"
            Call MySub(2,3)
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

    <div class="box">
        <h3>6. Array</h3>
    <%
    Dim famname(5),i
    famname(0) = "Jan Egil"
    famname(1) = "Tove"
    famname(2) = "Hege"
    famname(3) = "Stale"
    famname(4) = "Kai Jim"
    famname(5) = "Borge"
    
    For i = 0 to 5
         response.write(famname(i) & "<br>")
    Next
    %>
    </div>

    <div class="box">
        <h3>7. Replace variable</h3>
    <%
    Dim firstname
    firstname="Hege"
    response.write(firstname)
    response.write("<br>")
    firstname="Tove"
    response.write(firstname)
    %>
    </div>

    <div class="box">
        <h3>8. Request Object (QueryString & Form)</h3>
        <p>Test with ?msg=Hello in the URL.</p>
        <%
            Dim msg
            msg = Request.QueryString("msg")
            If msg <> "" Then
                Response.Write "<b>QueryString 'msg':</b> " & msg & "<br>"
            Else
                Response.Write "No 'msg' parameter in QueryString.<br>"
            End If
        %>
        <hr>
        <form method="POST">
            <input type="text" name="data" placeholder="Enter data">
            <input type="submit" value="Post Data">
        </form>
        <%
            Dim data
            data = Request.Form("data")
            If data <> "" Then
                Response.Write "<b>Received Form 'data':</b> " & data & "<br>"
            End If
        %>
    </div>

    </div>
</body>
</html>
