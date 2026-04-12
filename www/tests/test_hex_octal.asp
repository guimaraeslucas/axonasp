<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <title>Hexadecimal and Octal Literals Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #333; }
        .test { margin: 10px 0; padding: 10px; border: 1px solid #ddd; }
        .pass { background-color: #d4edda; }
        .fail { background-color: #f8d7da; }
        .value { font-weight: bold; color: #0066cc; }
    </style>
</head>
<body>
    <h1>G3pix AxonASP - Hexadecimal and Octal Literals Test</h1>
    
    <h2>Hexadecimal Literals (&amp;h prefix)</h2>
    
    <div class="test">
        <strong>Test 1:</strong> Basic hex literal &amp;h5C (should be 92)
        <% 
        cRevSolidus = &h5C
        Response.Write "<br/>Result: <span class='value'>" & cRevSolidus & "</span>"
        If cRevSolidus = 92 Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <div class="test">
        <strong>Test 2:</strong> Hex literal &amp;hFF (should be 255)
        <% 
        mask = &hFF
        Response.Write "<br/>Result: <span class='value'>" & mask & "</span>"
        If mask = 255 Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <div class="test">
        <strong>Test 3:</strong> Hex literal &amp;hFF00 (should be 65280)
        <% 
        bigMask = &hFF00
        Response.Write "<br/>Result: <span class='value'>" & bigMask & "</span>"
        If bigMask = 65280 Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <div class="test">
        <strong>Test 4:</strong> Lowercase hex &amp;h10 (should be 16)
        <% 
        lowerHex = &h10
        Response.Write "<br/>Result: <span class='value'>" & lowerHex & "</span>"
        If lowerHex = 16 Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <div class="test">
        <strong>Test 5:</strong> Uppercase hex &amp;hA0 (should be 160)
        <% 
        upperHex = &hA0
        Response.Write "<br/>Result: <span class='value'>" & upperHex & "</span>"
        If upperHex = 160 Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <div class="test">
        <strong>Test 6:</strong> Hex in arithmetic: &amp;h10 + &amp;h20 (should be 48)
        <% 
        sum = &h10 + &h20
        Response.Write "<br/>Result: <span class='value'>" & sum & "</span>"
        If sum = 48 Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <h2>Octal Literals (&amp;o prefix)</h2>
    
    <div class="test">
        <strong>Test 7:</strong> Basic octal literal &amp;o10 (should be 8)
        <% 
        octal1 = &o10
        Response.Write "<br/>Result: <span class='value'>" & octal1 & "</span>"
        If octal1 = 8 Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <div class="test">
        <strong>Test 8:</strong> Octal literal &amp;o77 (should be 63)
        <% 
        octal2 = &o77
        Response.Write "<br/>Result: <span class='value'>" & octal2 & "</span>"
        If octal2 = 63 Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <div class="test">
        <strong>Test 9:</strong> Octal literal &amp;o100 (should be 64)
        <% 
        octal3 = &o100
        Response.Write "<br/>Result: <span class='value'>" & octal3 & "</span>"
        If octal3 = 64 Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <div class="test">
        <strong>Test 10:</strong> Octal in arithmetic: &amp;o10 * &amp;o10 (should be 64)
        <% 
        octalMult = &o10 * &o10
        Response.Write "<br/>Result: <span class='value'>" & octalMult & "</span>"
        If octalMult = 64 Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <h2>String Concatenation with &amp; (should still work)</h2>
    
    <div class="test">
        <strong>Test 11:</strong> String concatenation "Hello" &amp; " World"
        <% 
        concat1 = "Hello" & " World"
        Response.Write "<br/>Result: <span class='value'>" & concat1 & "</span>"
        If concat1 = "Hello World" Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <div class="test">
        <strong>Test 12:</strong> Mixed concatenation: "Value is " &amp; 123
        <% 
        concat2 = "Value is " & 123
        Response.Write "<br/>Result: <span class='value'>" & concat2 & "</span>"
        If concat2 = "Value is 123" Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <h2>Complex Expressions</h2>
    
    <div class="test">
        <strong>Test 13:</strong> Mixed types: (&amp;hFF + &amp;o10) * 2 (should be 526)
        <% 
        complex1 = (&hFF + &o10) * 2
        Response.Write "<br/>Result: <span class='value'>" & complex1 & "</span>"
        If complex1 = 526 Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <div class="test">
        <strong>Test 14:</strong> Hex in string output: "Color: #" &amp; Hex(&amp;hFF)
        <% 
        colorStr = "Color: #" & Hex(&hFF)
        Response.Write "<br/>Result: <span class='value'>" & colorStr & "</span>"
        If colorStr = "Color: #FF" Then
            Response.Write " <span class='pass'>✓ PASS</span>"
        Else
            Response.Write " <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <div class="test">
        <strong>Test 15:</strong> Comparison: &amp;h64 = 100 (should be true)
        <% 
        Dim isEqual
        If &h64 = 100 Then
            isEqual = "True"
            Response.Write "<br/>Result: <span class='value'>" & isEqual & "</span> <span class='pass'>✓ PASS</span>"
        Else
            isEqual = "False"
            Response.Write "<br/>Result: <span class='value'>" & isEqual & "</span> <span class='fail'>✗ FAIL</span>"
        End If
        %>
    </div>
    
    <hr>
    <p><strong>Test Summary:</strong> All tests should display PASS for proper hex and octal literal support.</p>
    <p><a href="default.asp">← Back to Default Page</a></p>
</body>
</html>
