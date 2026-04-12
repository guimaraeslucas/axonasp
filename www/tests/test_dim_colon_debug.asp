<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Dim with Colon Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 25px; margin-bottom: 15px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .test { border-left: 4px solid #667eea; padding: 15px; margin: 15px 0; background: #f9f9f9; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Dim with Colon Test</h1>
        <div class="intro">
            <p>Tests Dim variable declaration with colon syntax on same line as initialization.</p>
        </div>
    
    <div class="test">
        <h2>Test 1: Simple case</h2>
        <% 
        Dim result1
        result1 = 42
        Response.Write "result1 = " & result1 & "<br/>"
        %>
    </div>
    
    <div class="test">
        <h2>Test 2: Same line with colon</h2>
        <% 
        Dim result2 : result2 = 42
        Response.Write "result2 = " & result2 & "<br/>"
        %>
    </div>
    
    <div class="test">
        <h2>Test 3: Function call without colon</h2>
        <% 
        Function triple(x)
            triple = x * 3
        End Function
        
        Dim result3
        result3 = triple(5)
        Response.Write "result3 = triple(5) = " & result3 & "<br/>"
        %>
    </div>
    
    <div class="test">
        <h2>Test 4: Function call WITH colon</h2>
        <% 
        Function quad(x)
            quad = x * 4
        End Function
        
        Dim result4 : result4 = quad(5)
        Response.Write "result4 = quad(5) = " & result4 & "<br/>"
        %>
    </div>
    </div>
</body>
</html>
