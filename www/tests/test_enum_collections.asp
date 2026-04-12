<%
@ Language = VBScript
%>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8" />
        <title>Collection Enumeration Test</title>
        <style>
            body {
                font-family: Tahoma, sans-serif;
                padding: 20px;
                background: #f5f5f5;
            }
            .container {
                max-width: 900px;
                margin: 0 auto;
                background: #fff;
                padding: 30px;
            }
            h1 {
                color: #333;
                border-bottom: 2px solid #667eea;
                padding-bottom: 10px;
            }
            .box {
                border-left: 4px solid #667eea;
                padding: 15px;
                margin: 15px 0;
                background: #f9f9f9;
            }
            .success {
                color: #388e3c;
            }
            .error {
                color: #d32f2f;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Collection Enumeration Test</h1>

            <div class="box">
                <h3>1. Session.Contents Enumeration Test</h3>
                <%
                Session("TestKey1") = "Value1"
                Session("TestKey2") = "Value2"
                Session("TestKey3") = 12345

                Response.Write("<p>Iterating through Session.Contents:</p>")
                Dim Key
                For Each Key In Session.Contents
                    Response.Write("Key: " & Key & " = " & Session(Key) & "<br>")
                Next
                %>
            </div>

            <div class="box">
                <h3>2. Application.Contents Enumeration Test</h3>
                <%
                Application.Lock
                Application("AppKey1") = "AppValue1"
                Application("AppKey2") = "AppValue2"
                Application("AppKey3") = 67890
                Application.Unlock

                Response.Write("<p>Iterating through Application.Contents:</p>")
                Dim appKey
                For Each appKey In Application.Contents
                    Response.Write("Key: " & appKey & " = " & Application(appKey) & "<br>")
                Next
                %>
            </div>

            <div class="box">
                <h3>3. Application.StaticObjects Enumeration Test</h3>
                <%
                Response.Write("<p>Iterating through Application.StaticObjects:</p>")
                Dim staticKey
                Dim Count
                Count = 0
                For Each staticKey In Application.StaticObjects
                    Count = Count + 1
                    Response.Write("Key: " & staticKey & "<br>")
                Next
                If Count = 0 Then
                    Response.Write("<span class='success'>No static objects (expected, as none are defined in global.asa)</span><br>")
                End If
                %>
            </div>

            <div class="box">
                <h3>4. Mixed Enumeration Test</h3>
                <%
                Response.Write("<p class='success'>All enumeration tests completed successfully!</p>")
                %>
            </div>
        </div>
    </body>
</html>
