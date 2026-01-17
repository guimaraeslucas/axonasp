<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Type Functions and Literals Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1200px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; border-left: 4px solid #667eea; padding-left: 15px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .test { background: #f9f9f9; border-left: 4px solid #667eea; padding: 15px; margin: 15px 0; border-radius: 4px; }
        .pass { color: #28a745; font-weight: bold; }
        .fail { color: #dc3545; font-weight: bold; }
        .value { font-weight: bold; color: #0066cc; }
        code { background: #eee; padding: 2px 6px; border-radius: 3px; font-family: 'Courier New', monospace; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Type Functions and Literals Test</h1>
        <div class="intro">
            <p>Tests TypeName, VarType, RGB, IsObject, IsEmpty, IsNull, Is Nothing operators, and hexadecimal/octal literals.</p>
        </div>

        <h2>Hexadecimal and Octal Literals</h2>
        <div class="test">
            <strong>Test 1:</strong> Hexadecimal literal &amp;h5C (should be 92)
            <%
            Dim hexValue
            hexValue = &h5C
            Response.Write "<br/>Result: <span class='value'>" & hexValue & "</span>"
            If hexValue = 92 Then
                Response.Write " <span class='pass'>✓ PASS</span>"
            Else
                Response.Write " <span class='fail'>✗ FAIL</span>"
            End If
            %>
        </div>

        <div class="test">
            <strong>Test 2:</strong> Hexadecimal literal &amp;hFF (should be 255)
            <%
            hexValue = &hFF
            Response.Write "<br/>Result: <span class='value'>" & hexValue & "</span>"
            If hexValue = 255 Then
                Response.Write " <span class='pass'>✓ PASS</span>"
            Else
                Response.Write " <span class='fail'>✗ FAIL</span>"
            End If
            %>
        </div>

        <div class="test">
            <strong>Test 3:</strong> Octal literal &amp;o10 (should be 8)
            <%
            Dim octValue
            octValue = &o10
            Response.Write "<br/>Result: <span class='value'>" & octValue & "</span>"
            If octValue = 8 Then
                Response.Write " <span class='pass'>✓ PASS</span>"
            Else
                Response.Write " <span class='fail'>✗ FAIL</span>"
            End If
            %>
        </div>

        <div class="test">
            <strong>Test 4:</strong> Octal literal &amp;o77 (should be 63)
            <%
            octValue = &o77
            Response.Write "<br/>Result: <span class='value'>" & octValue & "</span>"
            If octValue = 63 Then
                Response.Write " <span class='pass'>✓ PASS</span>"
            Else
                Response.Write " <span class='fail'>✗ FAIL</span>"
            End If
            %>
        </div>

        <h2>TypeName Function</h2>
        <div class="test">
            <%
            Dim testStr, testInt, testFloat, testBool, testEmpty, testNull, testNothing, testArr
            testStr = "Hello"
            testInt = 42
            testFloat = 3.14
            testBool = True
            testEmpty = Empty
            testNull = Null
            Set testNothing = Nothing
            testArr = Array(1, 2, 3)

            Response.Write "<strong>TypeName Tests:</strong><br/>"
            Response.Write "TypeName('Hello'): <span class='value'>" & TypeName(testStr) & "</span> (expected: String)<br/>"
            Response.Write "TypeName(42): <span class='value'>" & TypeName(testInt) & "</span> (expected: Integer)<br/>"
            Response.Write "TypeName(3.14): <span class='value'>" & TypeName(testFloat) & "</span> (expected: Double)<br/>"
            Response.Write "TypeName(True): <span class='value'>" & TypeName(testBool) & "</span> (expected: Boolean)<br/>"
            Response.Write "TypeName(Empty): <span class='value'>" & TypeName(testEmpty) & "</span> (expected: Empty)<br/>"
            Response.Write "TypeName(Null): <span class='value'>" & TypeName(testNull) & "</span> (expected: Empty)<br/>"
            Response.Write "TypeName(Nothing): <span class='value'>" & TypeName(testNothing) & "</span> (expected: Nothing)<br/>"
            Response.Write "TypeName(Array()): <span class='value'>" & TypeName(testArr) & "</span> (expected: Variant())<br/>"
            %>
        </div>

        <h2>VarType Function</h2>
        <div class="test">
            <%
            Response.Write "<strong>VarType Tests:</strong><br/>"
            Response.Write "VarType('Hello'): <span class='value'>" & VarType(testStr) & "</span> (expected: 8 = vbString)<br/>"
            Response.Write "VarType(42): <span class='value'>" & VarType(testInt) & "</span> (expected: 2 = vbInteger)<br/>"
            Response.Write "VarType(3.14): <span class='value'>" & VarType(testFloat) & "</span> (expected: 5 = vbDouble)<br/>"
            Response.Write "VarType(True): <span class='value'>" & VarType(testBool) & "</span> (expected: 11 = vbBoolean)<br/>"
            Response.Write "VarType(Empty): <span class='value'>" & VarType(testEmpty) & "</span> (expected: 0 = vbEmpty)<br/>"
            Response.Write "VarType(Array()): <span class='value'>" & VarType(testArr) & "</span> (expected: 8204 = vbArray + vbVariant)<br/>"
            %>
        </div>

        <h2>RGB Function</h2>
        <div class="test">
            <%
            Dim redColor, greenColor, blueColor, grayColor
            redColor = RGB(255, 0, 0)
            greenColor = RGB(0, 255, 0)
            blueColor = RGB(0, 0, 255)
            grayColor = RGB(128, 128, 128)

            Response.Write "<strong>RGB Tests:</strong><br/>"
            Response.Write "RGB(255, 0, 0) [Red]: <span class='value'>" & redColor & "</span> (expected: 255)<br/>"
            Response.Write "RGB(0, 255, 0) [Green]: <span class='value'>" & greenColor & "</span> (expected: 65280)<br/>"
            Response.Write "RGB(0, 0, 255) [Blue]: <span class='value'>" & blueColor & "</span> (expected: 16711680)<br/>"
            Response.Write "RGB(128, 128, 128) [Gray]: <span class='value'>" & grayColor & "</span> (expected: 8421504)<br/>"

            If redColor = 255 Then
                Response.Write "<span class='pass'>✓ Red PASS</span><br/>"
            Else
                Response.Write "<span class='fail'>✗ Red FAIL</span><br/>"
            End If
            %>
        </div>

        <h2>IsObject Function</h2>
        <div class="test">
            <%
            Dim simpleVar, objVar
            simpleVar = 42
            Set objVar = Server.CreateObject("Scripting.Dictionary")

            Response.Write "<strong>IsObject Tests:</strong><br/>"
            Response.Write "IsObject(42): <span class='value'>" & IsObject(simpleVar) & "</span> (expected: False)<br/>"
            Response.Write "IsObject('Hello'): <span class='value'>" & IsObject("Hello") & "</span> (expected: False)<br/>"
            Response.Write "IsObject(Dictionary): <span class='value'>" & IsObject(objVar) & "</span> (expected: True)<br/>"
            Response.Write "IsObject(Nothing): <span class='value'>" & IsObject(testNothing) & "</span> (expected: False)<br/>"
            %>
        </div>

        <h2>IsEmpty Function</h2>
        <div class="test">
            <%
            Dim newVar, assignedVar
            assignedVar = "Test"

            Response.Write "<strong>IsEmpty Tests:</strong><br/>"
            Response.Write "IsEmpty(newVar): <span class='value'>" & IsEmpty(newVar) & "</span> (expected: True)<br/>"
            Response.Write "IsEmpty(assignedVar): <span class='value'>" & IsEmpty(assignedVar) & "</span> (expected: False)<br/>"
            Response.Write "IsEmpty(Empty): <span class='value'>" & IsEmpty(Empty) & "</span> (expected: True)<br/>"
            
            If IsEmpty(newVar) And Not IsEmpty(assignedVar) Then
                Response.Write "<span class='pass'>✓ PASS</span><br/>"
            Else
                Response.Write "<span class='fail'>✗ FAIL</span><br/>"
            End If
            %>
        </div>

        <h2>IsNull Function</h2>
        <div class="test">
            <%
            Dim nullVar
            nullVar = Null

            Response.Write "<strong>IsNull Tests:</strong><br/>"
            Response.Write "IsNull(Null): <span class='value'>" & IsNull(nullVar) & "</span> (expected: True)<br/>"
            Response.Write "IsNull('Hello'): <span class='value'>" & IsNull("Hello") & "</span> (expected: False)<br/>"
            Response.Write "IsNull(42): <span class='value'>" & IsNull(42) & "</span> (expected: False)<br/>"
            
            If IsNull(nullVar) Then
                Response.Write "<span class='pass'>✓ PASS</span><br/>"
            Else
                Response.Write "<span class='fail'>✗ FAIL</span><br/>"
            End If
            %>
        </div>

        <h2>Is Nothing Operator</h2>
        <div class="test">
            <%
            Dim nothingVar, objVar2
            Set nothingVar = Nothing
            Set objVar2 = Server.CreateObject("Scripting.Dictionary")

            Response.Write "<strong>Is Nothing Tests:</strong><br/>"
            Response.Write "nothingVar Is Nothing: <span class='value'>" & (nothingVar Is Nothing) & "</span> (expected: True)<br/>"
            Response.Write "objVar2 Is Nothing: <span class='value'>" & (objVar2 Is Nothing) & "</span> (expected: False)<br/>"
            
            If nothingVar Is Nothing Then
                Response.Write "<span class='pass'>✓ Is Nothing PASS</span><br/>"
            Else
                Response.Write "<span class='fail'>✗ Is Nothing FAIL</span><br/>"
            End If
            
            If Not (objVar2 Is Nothing) Then
                Response.Write "<span class='pass'>✓ Is Not Nothing PASS</span><br/>"
            Else
                Response.Write "<span class='fail'>✗ Is Not Nothing FAIL</span><br/>"
            End If
            %>
        </div>

        <h2>Combined Tests</h2>
        <div class="test">
            <%
            Response.Write "<strong>Combined Literal and Function Tests:</strong><br/>"
            
            ' Test hex in arithmetic
            Dim hexSum
            hexSum = &h10 + &h20
            Response.Write "Hex arithmetic (&h10 + &h20): <span class='value'>" & hexSum & "</span> (expected: 48)"
            If hexSum = 48 Then
                Response.Write " <span class='pass'>✓ PASS</span><br/>"
            Else
                Response.Write " <span class='fail'>✗ FAIL</span><br/>"
            End If
            
            ' Test octal in arithmetic
            Dim octSum
            octSum = &o10 + &o20
            Response.Write "Octal arithmetic (&o10 + &o20): <span class='value'>" & octSum & "</span> (expected: 24)"
            If octSum = 24 Then
                Response.Write " <span class='pass'>✓ PASS</span><br/>"
            Else
                Response.Write " <span class='fail'>✗ FAIL</span><br/>"
            End If
            
            ' Test TypeName with hex value
            Response.Write "TypeName(&hFF): <span class='value'>" & TypeName(&hFF) & "</span> (expected: Integer)<br/>"
            
            ' Test VarType with octal value
            Response.Write "VarType(&o77): <span class='value'>" & VarType(&o77) & "</span> (expected: 2 = vbInteger)<br/>"
            %>
        </div>

    </div>
</body>
</html>
