<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Control Flow Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .result { background: #f9f9f9; border-left: 4px solid #667eea; padding: 15px; margin: 15px 0; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Control Flow Test</h1>
        <div class="intro">
            <p>Tests For/Next loops with Step, Do/Loop, Select/Case and If/ElseIf/Else.</p>
        </div>

<h2>For ... Step</h2>
<%
    Dim i
    Response.Write "Count 1 to 10 step 2: "
    For i = 1 To 10 Step 2
        Response.Write i & " "
    Next
    Response.Write "<br>"
    
    Response.Write "Count 10 to 1 step -2: "
    For i = 10 To 1 Step -2
        Response.Write i & " "
    Next
    Response.Write "<br>"
%>

<h2>Do ... Loop</h2>
<%
    Dim x
    x = 1
    Response.Write "Do Until x > 5: "
    Do Until x > 5
        Response.Write x & " "
        x = x + 1
    Loop
    Response.Write "<br>"
    
    x = 1
    Response.Write "Do While x <= 5 (using Exit Do at 3): "
    Do While x <= 5
        If x = 3 Then Exit Do
        Response.Write x & " "
        x = x + 1
    Loop
    Response.Write "<br>"
%>

<h2>Select Case</h2>
<%
    Dim color
    color = "Blue"
    Response.Write "Select Case '" & color & "': "
    
    Select Case color
        Case "Red"
            Response.Write "It is Red."
        Case "Green"
            Response.Write "It is Green."
        Case "Blue"
            Response.Write "It is Blue."
        Case Else
            Response.Write "Unknown Color."
    End Select
    Response.Write "<br>"
%>

<h2>If ... ElseIf ... Else</h2>
<%
    Dim score
    score = 75
    Response.Write "Score " & score & ": "
    
    If score >= 90 Then
        Response.Write "Grade A"
    ElseIf score >= 80 Then
        Response.Write "Grade B"
    ElseIf score >= 70 Then
        Response.Write "Grade C"
    Else
        Response.Write "Grade F"
    End If
    Response.Write "<br>"
%>
    </div>

    </div>
</body>
</html>

