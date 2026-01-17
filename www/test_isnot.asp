<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Is Not & Document.WriteSafe Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3 { color: #666; margin-top: 20px; margin-bottom: 10px; }
        p { margin: 10px 0; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Is Not & Document.WriteSafe Test</h1>
        <div class="intro">
            <p>Tests the Is Not operator for object comparison and Document.Write vs Document.WriteSafe methods.</p>
        </div>

    <h3>1. Is Not Nothing Check</h3>
    <%
        Dim obj
        ' obj is Empty initially
        
        If obj Is Nothing Then
            Response.Write "Case 1: obj is Nothing (Correct for Empty?)<br>"
        Else
            Response.Write "Case 1: obj is Not Nothing (Correct for Empty)<br>"
        End If

        Set obj = Nothing
        If obj Is Nothing Then
             Response.Write "Case 2: obj is Nothing (Correct)<br>"
        Else
             Response.Write "Case 2: Fail<br>"
        End If

        If obj Is Not Nothing Then
             Response.Write "Case 3: Fail<br>"
        Else
             Response.Write "Case 3: obj Is Not Nothing is False (Correct)<br>"
        End If

        Set obj = JSON.NewObject()
        If obj Is Nothing Then
             Response.Write "Case 4: Fail<br>"
        End If
        
        If obj Is Not Nothing Then
             Response.Write "Case 4: obj Is Not Nothing (Correct)<br>"
        Else
             Response.Write "Case 4: Fail (Thought it was Nothing)<br>"
        End If
    %>

    <h3>2. Document.Write (Raw) vs Document.WriteSafe (Encoded)</h3>
    <p>Document.Write (Raw): Should show bold text.</p>
    <%
        Document.Write "<b>Bold Text</b>"
    %>
    
    <p>Document.WriteSafe (Encoded): Should show the HTML code literal.</p>
    <%
        Document.WriteSafe "<b>This HTML code is visible, not rendered</b>"
    %>
    </div>
</body>
</html>