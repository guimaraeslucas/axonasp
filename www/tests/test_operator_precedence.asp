<%
Response.Write "<h2>Operator Precedence Test</h2>"

Dim result

' Should evaluate as: 2 + (3 * 4) = 14
result = 2 + 3 * 4
Response.Write "2 + 3 * 4 = " & result & " (expected 14)<br>"

' Should evaluate as: (5 And 4) Or 2 = 4 Or 2 = 6
' VBScript precedence: And has higher precedence than Or
result = 5 And 4 Or 2
Response.Write "5 And 4 Or 2 = " & result & " (expected 6)<br>"

' Exponentiation has higher precedence than unary minus: -(2^2) = -4
result = - 2 ^ 2
Response.Write "-2 ^ 2 = " & result & " (expected -4)<br>"

' Parentheses force unary before exponentiation: (-2)^2 = 4
result = ( - 2) ^ 2
Response.Write "(-2) ^ 2 = " & result & " (expected 4)<br>"

' Right-associative exponentiation: 2^(3^2) = 512
result = 2 ^ 3 ^ 2
Response.Write "2 ^ 3 ^ 2 = " & result & " (expected 512)<br>"

' Signed exponent RHS: 2^(-2) = 0.25
result = 2 ^ - 2
Response.Write "2 ^ -2 = " & result & " (expected 0.25)<br>"

' Not has lower precedence than comparison operators
result = Not 1 = 1
Response.Write "Not 1 = 1 -> " & result & " (expected False)<br>"

%>
