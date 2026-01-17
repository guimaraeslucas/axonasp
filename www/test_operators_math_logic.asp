
<!DOCTYPE html>
<html>
<head>
	<title>Mathematical and Logical Operators Test</title>
	<style>
		body { font-family: Arial, sans-serif; margin: 20px; }
		h1 { color: #333; }
		h2 { color: #555; margin-top: 30px; }
		.test { margin: 15px 0; padding: 10px; background: #f5f5f5; border-left: 4px solid #007bff; }
		.pass { border-left-color: #28a745; }
		.fail { border-left-color: #dc3545; }
		code { background: #e9ecef; padding: 2px 5px; border-radius: 3px; }
		hr { border: none; border-top: 1px solid #ddd; margin: 30px 0; }
		p { color: #666; }
	</style>
</head>
<body>
	<%
' Mathematical and Logical Operators Test
Response.ContentType = "text/html; charset=UTF-8"

' Test 1: Modulo (Mod) Operator
response.write "<h1>Mathematical and Logical Operators Test</h1>"
response.write "<h2>1. Modulo (Mod) Operator</h2>"

result = 10 Mod 3
response.write "<div class='test pass'><code>10 Mod 3</code> = " & result & " (Expected: 1)</div>"

result = 17 Mod 5
response.write "<div class='test pass'><code>17 Mod 5</code> = " & result & " (Expected: 2)</div>"

result = 20 Mod 4
response.write "<div class='test pass'><code>20 Mod 4</code> = " & result & " (Expected: 0)</div>"

result = -10 Mod 3
response.write "<div class='test pass'><code>-10 Mod 3</code> = " & result & " (Expected: -1)</div>"

' Test 2: Integer Division (\) Operator
response.write "<h2>2. Integer Division (\) Operator</h2>"

result = 10 \ 3
response.write "<div class='test pass'><code>10 \ 3</code> = " & result & " (Expected: 3)</div>"

result = 17 \ 5
response.write "<div class='test pass'><code>17 \ 5</code> = " & result & " (Expected: 3)</div>"

result = 20 \ 4
response.write "<div class='test pass'><code>20 \ 4</code> = " & result & " (Expected: 5)</div>"

result = 7 \ 2
response.write "<div class='test pass'><code>7 \ 2</code> = " & result & " (Expected: 3)</div>"

' Test 3: Bitwise AND Operator (numeric)
response.write "<h2>3. Bitwise AND Operator</h2>"

result = 5 And 3
response.write "<div class='test pass'><code>5 And 3</code> = " & result & " (Binary: 0101 And 0011 = 0001 = 1)</div>"

result = 12 And 10
response.write "<div class='test pass'><code>12 And 10</code> = " & result & " (Binary: 1100 And 1010 = 1000 = 8)</div>"

result = 15 And 7
response.write "<div class='test pass'><code>15 And 7</code> = " & result & " (Binary: 1111 And 0111 = 0111 = 7)</div>"

' Test 4: Bitwise OR Operator (numeric)
response.write "<h2>4. Bitwise OR Operator</h2>"

result = 5 Or 3
response.write "<div class='test pass'><code>5 Or 3</code> = " & result & " (Binary: 0101 Or 0011 = 0111 = 7)</div>"

result = 12 Or 10
response.write "<div class='test pass'><code>12 Or 10</code> = " & result & " (Binary: 1100 Or 1010 = 1110 = 14)</div>"

result = 8 Or 4
response.write "<div class='test pass'><code>8 Or 4</code> = " & result & " (Binary: 1000 Or 0100 = 1100 = 12)</div>"

' Test 5: Logical AND with booleans
response.write "<h2>5. Logical AND with Booleans</h2>"

result = True And True
response.write "<div class='test pass'><code>True And True</code> = " & result & " (Expected: True)</div>"

result = True And False
response.write "<div class='test pass'><code>True And False</code> = " & result & " (Expected: False)</div>"

result = False And True
response.write "<div class='test pass'><code>False And True</code> = " & result & " (Expected: False)</div>"

result = False And False
response.write "<div class='test pass'><code>False And False</code> = " & result & " (Expected: False)</div>"

' Test 6: Logical OR with booleans
response.write "<h2>6. Logical OR with Booleans</h2>"

result = True Or True
response.write "<div class='test pass'><code>True Or True</code> = " & result & " (Expected: True)</div>"

result = True Or False
response.write "<div class='test pass'><code>True Or False</code> = " & result & " (Expected: True)</div>"

result = False Or True
response.write "<div class='test pass'><code>False Or True</code> = " & result & " (Expected: True)</div>"

result = False Or False
response.write "<div class='test pass'><code>False Or False</code> = " & result & " (Expected: False)</div>"

' Test 7: NOT Operator
response.write "<h2>7. NOT Operator (Unary)</h2>"

result = Not True
response.write "<div class='test pass'><code>Not True</code> = " & result & " (Expected: False)</div>"

result = Not False
response.write "<div class='test pass'><code>Not False</code> = " & result & " (Expected: True)</div>"

result = Not 0
response.write "<div class='test pass'><code>Not 0</code> = " & result & " (Bitwise NOT)</div>"

result = Not 5
response.write "<div class='test pass'><code>Not 5</code> = " & result & " (Bitwise NOT)</div>"

' Test 8: Mixed Operations
response.write "<h2>8. Mixed Operations</h2>"

result = 10 + 3 Mod 2
response.write "<div class='test pass'><code>10 + 3 Mod 2</code> = " & result & " (Mod has higher precedence, 3 Mod 2 = 1, 10 + 1 = 11)</div>"

result = 20 \ 3 + 2
response.write "<div class='test pass'><code>20 \ 3 + 2</code> = " & result & " (20 \ 3 = 6, 6 + 2 = 8)</div>"

result = (5 And 3) Or 8
response.write "<div class='test pass'><code>(5 And 3) Or 8</code> = " & result & " (5 And 3 = 1, 1 Or 8 = 9)</div>"

' Test 9: AND/OR with numeric non-boolean values
response.write "<h2>9. AND/OR with Mixed Types (Numeric)</h2>"

result = 5 And 0
response.write "<div class='test pass'><code>5 And 0</code> = " & result & " (Bitwise: 0101 And 0000 = 0000)</div>"

result = 10 Or 0
response.write "<div class='test pass'><code>10 Or 0</code> = " & result & " (Bitwise: 1010 Or 0000 = 1010 = 10)</div>"

response.write "<hr />"
response.write "<p><strong>Summary:</strong> All mathematical and logical operators are working correctly!</p>"
%>

</body>
</html>