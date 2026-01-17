<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3 AxonASP - Operators Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .result { background: #f9f9f9; border-left: 4px solid #667eea; padding: 15px; margin: 15px 0; border-radius: 4px; }
        .success { background: #e8f5e9; border-left: 4px solid #4caf50; }
        .error { background: #ffebee; border-left: 4px solid #f44336; }
        code { background: #f5f5f5; padding: 2px 6px; border-radius: 3px; font-family: 'Courier New', monospace; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3 AxonASP - Operators Test</h1>
        <div class="intro">
            <p>Tests mathematical and logical operators: <code>Mod</code>, <code>\</code>, <code>And</code>, <code>Or</code>, <code>Not</code>, <code>+</code>, <code>-</code>, <code>*</code>, <code>/</code>, and special values <code>Null</code>, <code>Empty</code>, <code>Nothing</code>, <code>True</code>, <code>False</code>.</p>
        </div>

        <h2>Arithmetic Operators</h2>
        <div class="result success">
            <%
                Dim a, b, result
                a = 10
                b = 3
                
                Response.Write "<p><strong>Addition (+):</strong> " & a & " + " & b & " = " & (a + b) & "</p>"
                Response.Write "<p><strong>Subtraction (-):</strong> " & a & " - " & b & " = " & (a - b) & "</p>"
                Response.Write "<p><strong>Multiplication (*):</strong> " & a & " * " & b & " = " & (a * b) & "</p>"
                Response.Write "<p><strong>Division (/):</strong> " & a & " / " & b & " = " & (a / b) & "</p>"
            %>
        </div>

        <h2>Integer Division and Modulo</h2>
        <div class="result success">
            <%
                a = 17
                b = 5
                
                Response.Write "<p><strong>Integer Division (\):</strong> " & a & " \ " & b & " = " & (a \ b) & " (should be 3)</p>"
                Response.Write "<p><strong>Modulo (Mod):</strong> " & a & " Mod " & b & " = " & (a Mod b) & " (should be 2)</p>"
            %>
        </div>

        <h2>Logical Operators</h2>
        <div class="result success">
            <%
                Dim x, y
                x = True
                y = False
                
                Response.Write "<p><strong>And:</strong> True And False = " & (x And y) & " (should be False)</p>"
                Response.Write "<p><strong>Or:</strong> True Or False = " & (x Or y) & " (should be True)</p>"
                Response.Write "<p><strong>Not:</strong> Not True = " & (Not x) & " (should be False)</p>"
                Response.Write "<p><strong>Not:</strong> Not False = " & (Not y) & " (should be True)</p>"
            %>
        </div>

        <h2>Bitwise Operations</h2>
        <div class="result success">
            <%
                Dim val1, val2
                val1 = 12  ' Binary: 1100
                val2 = 10  ' Binary: 1010
                
                Response.Write "<p><strong>Bitwise And:</strong> " & val1 & " And " & val2 & " = " & (val1 And val2) & " (should be 8, binary 1000)</p>"
                Response.Write "<p><strong>Bitwise Or:</strong> " & val1 & " Or " & val2 & " = " & (val1 Or val2) & " (should be 14, binary 1110)</p>"
                Response.Write "<p><strong>Bitwise Not:</strong> Not " & val1 & " = " & (Not val1) & " (inverts all bits)</p>"
            %>
        </div>

        <h2>Boolean Values</h2>
        <div class="result success">
            <%
                Dim boolTrue, boolFalse
                boolTrue = True
                boolFalse = False
                
                Response.Write "<p><strong>True value:</strong> " & boolTrue & " (numeric value: " & CLng(boolTrue) & ")</p>"
                Response.Write "<p><strong>False value:</strong> " & boolFalse & " (numeric value: " & CLng(boolFalse) & ")</p>"
                Response.Write "<p><strong>True = -1:</strong> " & (boolTrue = -1) & "</p>"
                Response.Write "<p><strong>False = 0:</strong> " & (boolFalse = 0) & "</p>"
            %>
        </div>

        <h2>Special Values</h2>
        <div class="result success">
            <%
                Dim varNull, varEmpty, varNothing
                varEmpty = Empty
                varNull = Null
                Set varNothing = Nothing
                
                Response.Write "<p><strong>Empty:</strong> IsEmpty(varEmpty) = " & IsEmpty(varEmpty) & " (should be True)</p>"
                Response.Write "<p><strong>Null:</strong> IsNull(varNull) = " & IsNull(varNull) & " (should be True)</p>"
                Response.Write "<p><strong>Nothing:</strong> varNothing Is Nothing = " & (varNothing Is Nothing) & " (should be True)</p>"
            %>
        </div>

        <h2>Operator Precedence</h2>
        <div class="result success">
            <%
                Response.Write "<p><strong>Expression:</strong> 2 + 3 * 4 = " & (2 + 3 * 4) & " (should be 14)</p>"
                Response.Write "<p><strong>Expression:</strong> (2 + 3) * 4 = " & ((2 + 3) * 4) & " (should be 20)</p>"
                Response.Write "<p><strong>Expression:</strong> 10 - 3 + 2 = " & (10 - 3 + 2) & " (should be 9)</p>"
                Response.Write "<p><strong>Expression:</strong> 15 Mod 4 * 2 = " & (15 Mod 4 * 2) & " (should be 6)</p>"
            %>
        </div>

        <h2>Comparison with Logical Operators</h2>
        <div class="result success">
            <%
                a = 5
                b = 10
                Dim c
                c = 15
                
                Response.Write "<p><strong>Expression:</strong> (a < b) And (b < c) = " & ((a < b) And (b < c)) & " (should be True)</p>"
                Response.Write "<p><strong>Expression:</strong> (a > b) Or (b < c) = " & ((a > b) Or (b < c)) & " (should be True)</p>"
                Response.Write "<p><strong>Expression:</strong> Not (a = b) = " & (Not (a = b)) & " (should be True)</p>"
            %>
        </div>

        <h2>Negative Numbers</h2>
        <div class="result success">
            <%
                Dim neg1, neg2
                neg1 = -10
                neg2 = -5
                
                Response.Write "<p><strong>Addition:</strong> " & neg1 & " + " & neg2 & " = " & (neg1 + neg2) & " (should be -15)</p>"
                Response.Write "<p><strong>Multiplication:</strong> " & neg1 & " * " & neg2 & " = " & (neg1 * neg2) & " (should be 50)</p>"
                Response.Write "<p><strong>Modulo with negative:</strong> " & neg1 & " Mod 3 = " & (neg1 Mod 3) & "</p>"
            %>
        </div>
    </div>
</body>
</html>
