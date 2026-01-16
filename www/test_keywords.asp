<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Keywords & Types Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .result { border-left: 4px solid #667eea; padding: 15px; margin: 10px 0; background: #f9f9f9; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Keywords & Types Test</h1>
        <div class="intro">
            <p>Tests VBScript keywords like Empty, Null, Nothing, True, False and verification functions.</p>
        </div>
        <div class="result">
<% Response.Write("<h2>Teste de Keywords VBScript</h2>")

' Test 1: True and False keywords
Response.Write("<h3>Test 1: True and False</h3>")
If True Then
    Response.Write("True keyword works: " & True & "<br>")
End If
If Not False Then
    Response.Write("False keyword works: " & False & "<br>")
End If

' Test 2: Empty keyword
Response.Write("<h3>Test 2: Empty</h3>")
Dim emptyVar
emptyVar = Empty
Response.Write("Empty value: " & emptyVar & "<br>")
If IsEmpty(emptyVar) Then
    Response.Write("IsEmpty() correctly identifies Empty<br>")
End If

' Test 3: Null keyword
Response.Write("<h3>Test 3: Null</h3>")
Dim nullVar
nullVar = Null
Response.Write("Null value test<br>")
If IsNull(nullVar) Then
    Response.Write("IsNull() correctly identifies Null<br>")
End If

' Test 4: Nothing keyword
Response.Write("<h3>Test 4: Nothing</h3>")
Dim objVar
objVar = Nothing
Response.Write("Nothing value test<br>")
If objVar Is Nothing Then
    Response.Write("Is Nothing operator correctly identifies Nothing<br>")
End If

' Test 5: Is Not Nothing
Response.Write("<h3>Test 5: Is Not Nothing</h3>")
Dim someVar
someVar = "test"
If Not (someVar Is Nothing) Then
    Response.Write("Is Not Nothing operator works correctly<br>")
End If

Response.Write("<h3>Testes de Keywords Completados!</h3>")
%>
        </div>
    </div>
</body>
</html>

